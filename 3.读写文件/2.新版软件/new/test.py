import numpy as np
import time
import pyautogui
from pylsl import StreamInlet, resolve_stream

# 设置频率和闪烁时间
frequencies = [10, 15, 20, 25]
duration = 5

# 初始化连接到LSL流的接口
print("正在连接到LSL流...")
streams = resolve_stream('type', 'EEG')
inlet = StreamInlet(streams[0])
print("连接成功！")

# 定义函数，以便在每个闪烁频率上闪烁
def flicker(freq):
    period = 1.0 / freq
    samples_per_period = int(period * 256)
    samples = np.zeros((samples_per_period, 1))
    samples[-10:] = 1.0 # 闪烁的持续时间为10个样本
    while True:
        for i in range(samples_per_period):
            yield samples[i]

# 开始闪烁
flicker_generators = [flicker(f) for f in frequencies]
flicker_states = [0 for _ in frequencies]
last_sample_time = time.monotonic()

while True:
    # 从LSL流中读取数据
    sample, _ = inlet.pull_sample()
    now = time.monotonic()
    delta = now - last_sample_time

    # 更新闪烁状态
    for i, flicker_state in enumerate(flicker_states):
        flicker_states[i] = next(flicker_generators[i]) if delta > 1.0 / (2*frequencies[i]) else flicker_state

    # 判断哪个闪烁频率被看到了
    detected = None
    for i, flicker_state in enumerate(flicker_states):
        if flicker_state > 0:
            if detected is None:
                detected = i
            else:
                detected = None
                break

    # 如果检测到了频率，则执行相应的动作
    if detected is not None:
        print("检测到频率：", frequencies[detected])
        if frequencies[detected] == 10:
            pyautogui.press('left')
        elif frequencies[detected] == 15:
            pyautogui.press('right')
        elif frequencies[detected] == 20:
            pyautogui.press('up')
        elif frequencies[detected] == 25:
            pyautogui.press('down')

    # 检查持续时间是否已到达
    if now - last_sample_time > duration:
        break

    # 等待下一个样本
    time.sleep(0.001)
