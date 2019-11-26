import tkinter as tk
from tkinter import filedialog
from pandas import pandas as pd
from numpy import array


root = tk.Tk(className="工作单")
root.geometry('600x400')


# 源文件
def get_file_name():
    source_path = filedialog.askopenfilename(title=u'选择文件')
    t.insert('insert', source_path)


# 保存到
def save_file():
    out_path = filedialog.asksaveasfilename(title=u'保存文件')
    t2.insert('insert', out_path)


# 执行excel操作
def run_excel():
    label["text"] = "数据处理中......"
    path = t.get()
    # path = "E:\\workSheet.xls"
    savePath = t2.get()
    # savePath = "E:\\workSheet11.xls"
    wb = pd.read_excel(path, sheet_name='Sheet1')  # 读取excel
    wb = delete_by_word(wb)
    wb = delete_1_2(wb)
    wb = delete_1_3(wb)
    wb = delete_1_4(wb)
    wb = delete_1_2_3(wb)
    wb.to_excel(savePath, sheet_name='Sheet1', index=False, header=True)
    label["text"] = "success"


# 根据关键词删除
def delete_by_word(wb):
    # data = wb.iloc[:, 1].values  # 读取需要做筛选的列的数据
    data = wb.loc[:, '错误类型'].values  # 读取需要做筛选的列的数据
    # 查找有此文本“附图标记全部缺失”的行
    data = wb[(data == '附图标记全部缺失')]  # 筛选出需要的数据
    arr = array(data.index)  # 转为数组
    # 删除 axis=0 删除行 =1删除列
    wb.drop(arr, axis=0, inplace=True)

    # data = wb.iloc[:, 1].values  # 读取需要做筛选的列的数据
    data = wb.loc[:, '错误类型'].values  # 读取需要做筛选的列的数据
    # 查找有此文本“关键词没有体现核心方案主题名称”的行
    data = wb[(data == '关键词没有体现核心方案主题名称')]  # 筛选出需要的数据
    arr = array(data.index)  # 转为数组
    # 删除 axis=0 删除行 =1删除列
    wb.drop(arr, axis=0, inplace=True)

    # data = wb.iloc[:, 1].values  # 读取需要做筛选的列的数据
    data = wb.loc[:, '错误类型'].values  # 读取需要做筛选的列的数据
    # 查找有此文本“名称没有体现核心方案对应技术主题”的行
    data = wb[(data == '名称没有体现核心方案对应技术主题')]  # 筛选出需要的数据
    arr = array(data.index)  # 转为数组
    # 删除 axis=0 删除行 =1删除列
    wb.drop(arr, axis=0, inplace=True)

    # data = wb.iloc[:, 1].values  # 读取需要做筛选的列的数据
    data = wb.loc[:, '错误类型'].values  # 读取需要做筛选的列的数据
    # 查找有此文本“其他技术方案中的发明信息中缺失技术主题”的行
    data = wb[(data == '其他技术方案中的发明信息中缺失技术主题')]  # 筛选出需要的数据
    arr = array(data.index)  # 转为数组
    # 删除 axis=0 删除行 =1删除列
    wb.drop(arr, axis=0, inplace=True)
    return wb


'''
    发明名称与原始名称简单重复
    仅需审核名称长度小于等于18的，其他无需审核，直接删除。
'''
def delete_1_4(wb):
    data = wb.loc[:, '错误类型'].values  # 读取需要做筛选的列的数据
    # data = wb.iloc[:, 1].values  # 读取需要做筛选的列的数据
    data = wb[(data == '发明名称与原始名称简单重复')]  # 筛选出需要的数据
    arrWord = data.get("备注2")  # 需要筛选的数据列
    arrIndex = array(data.index)  # 数据列对应的excel index

    for index, value in enumerate(arrWord):
        value = value.replace("名称长度为：", "")
        if int(value) >= 19:
            wb.drop(arrIndex[index], axis=0, inplace=True)  # 按照excel的索引删除
    return wb


def delete_1_2(wb):
    data = wb.loc[:, '错误类型'].values  # 读取需要做筛选的列的数据
    data = wb[(data == '存在不宜标引的关键词')]  # 筛选出需要的数据
    arrWord = data.get("具体说明")  # 需要筛选的数据列
    arrIndex = array(data.index)  # 数据列对应的excel index
    for index, value in enumerate(arrWord):
        value = value.replace("关键词：", "")
        if value == '稳定性' or value == '程序':
            wb.drop(arrIndex[index], axis=0, inplace=True)  # 按照excel的索引删除

    data = wb.loc[:, '错误类型'].values  # 读取需要做筛选的列的数据
    data = wb[(data == '名称未增加限定特征')]  # 筛选出需要的数据
    arrWord = data.get("具体说明")  # 需要筛选的数据列
    arrIndex = array(data.index)  # 数据列对应的excel index
    for index, value in enumerate(arrWord):
        value = value.replace("名称字数为", "")
        if int(value) >= 13:
            wb.drop(arrIndex[index], axis=0, inplace=True)  # 13个及以上的直接删除

    data = wb.loc[:, '错误类型'].values  # 读取需要做筛选的列的数据
    data = wb[(data == '摘要中使用了不必要的措词')]  # 筛选出需要的数据
    arrWord = data.get("具体说明")  # 需要筛选的数据列
    arrIndex = array(data.index)  # 数据列对应的excel index
    for index, value in enumerate(arrWord):
        value = value.replace("具体摘要不宜词汇：", "")
        if value == '新的' or value == '特征' or value == '板机' or value == '拌均' or value == '蒸气' or value == '疏基' or value == '齿合' or value == '淬灭' or value == '成份' or value == '泄露' or value == '罗纹' or value == '咀':
            wb.drop(arrIndex[index], axis=0, inplace=True)  # 按照excel的索引删除
    return wb


def delete_1_3(wb):
    data = wb.loc[:, '错误类型'].values  # 读取需要做筛选的列的数据
    # data = wb.iloc[:, 1].values  # 读取需要做筛选的列的数据
    data = wb[(data == '存在未规范化处理的关键词')]  # 筛选出需要的数据
    arrWord = data.get("备注1")  # 需要筛选的数据列
    arrIndex = array(data.index)  # 数据列对应的excel index

    for index, value in enumerate(arrWord):
        if value.find('NULL') == -1:
            wb.drop(arrIndex[index], axis=0, inplace=True)  # 按照excel的索引删除

    data = wb.loc[:, '错误类型'].values  # 读取需要做筛选的列的数据
    data = wb[(data == '自由标引的关键词字数过多')]  # 筛选出需要的数据
    arrWord = data.get("备注1")  # 需要筛选的数据列
    arrIndex = array(data.index)  # 数据列对应的excel index

    for index, value in enumerate(arrWord):
        value = value.replace("实际字符数：", "")
        if int(value) <= 11:
            wb.drop(arrIndex[index], axis=0, inplace=True)  # 实际字符数11个及以下直接删除
    return wb


def delete_1_2_3(wb):
    data = wb.loc[:, '错误类型'].values  # 读取需要做筛选的列的数据
    data = wb[(data == '核心方案与独立权利要求1的内容简单重复')]  # 筛选出需要的数据
    arr2Word = data.get("具体说明")  # 需要筛选的数据列
    arr2Index = array(data.index)  # 数据列对应的excel index
    for index, value in enumerate(arr2Word):
        value = value.replace("%", "")
        if int(value) < 100:
            wb.drop(arr2Index[index], axis=0, inplace=True)  # 按照excel的索引删除【具体说明】小于“100%”的

    # 上一步已经删除了一些，需要更新data
    data = wb.loc[:, '错误类型'].values  # 读取需要做筛选的列的数据
    data = wb[(data == '核心方案与独立权利要求1的内容简单重复')]  # 筛选出需要的数据
    arr3Word = data.get("备注1")  # 需要筛选的数据列
    arr3Index = array(data.index)  # 数据列对应的excel index
    for index, value in enumerate(arr3Word):
        if int(value) > 460:
            wb.drop(arr3Index[index], axis=0, inplace=True)  # 按照excel的索引删除【备注1】大于460的
    return wb


'''
定义界面
'''
b = tk.Button(root, text='数据源文件', width=10, height=2, command=get_file_name).pack()
t = tk.Entry(borderwidth=3, width=50)
t.pack()
b2 = tk.Button(root, text='另存为', width=10, height=2, command=save_file).pack()
t2 = tk.Entry(borderwidth=3, width=50)
t2.pack()
submitButton = tk.Button(root, text='开始', width=10, height=2, command=run_excel).pack()
label = tk.Label(root)  # text为显示的文本内容
label.pack()

root.mainloop()

# 用来单元测试
# if __name__ == "__main__":
#     run_excel()


# https://blog.csdn.net/CSU_GUO_LIANG/article/details/102772294
# https://www.cnblogs.com/bianzhiwei/p/11214994.html
# http://www.pythonheidong.com/blog/article/147319/
# Python3 pandas自定义输出excel样式