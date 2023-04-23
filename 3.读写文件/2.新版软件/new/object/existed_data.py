def ifexistedlist(list,index):
    if index < len(list):
        return [list[index]]
    else:
        return []
        
#给解析用的，需要拼接
def ifexistedlist2(list,index,type = 1):
    if list[index] != "":
        if type == 1:
            leftstr = "<div><p>"
            rightstr ="</p></div>"
            return [leftstr + list[index] + rightstr]
        if type == 2:
            leftstr = "<div><img>"
            rightstr = "</img></div>"
            return [leftstr + list[index] + rightstr]
    else:
        return []