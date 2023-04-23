def ifexistedlist(list,index):
    if index < len(list):
        return [list[index]]
    else:
        return []
        
#给解析用的，需要拼接
def ifexistedlist2(list,index):
    if list[index] != "":
        return [list[index]]

    else:
        return []