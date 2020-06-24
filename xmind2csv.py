# -*- coding: UTF-8 -*-
# Date: 2019-08-07
# Author: fuzhuang.zhou


import sys
from xmindparser import xmind_to_dict
import xlwt
import time

path_list = list()
note_list = list()
mark_map = {"priority-1": "P0", "priority-2": "P1", "priority-3": "P2"}

# import xmind
#
#
# def get_xmind_data(file_name):
#     workbook = xmind.load(file_name)
#     sheet = workbook.getPrimarySheet()
#     try:
#         print sheet.getRootTopic().getMarkers()[0].getMarkerId()
#     except Exception, e:
#         marker = None
#         print marker
#
#     print sheet.getRootTopic().getTitle()


def get_original_data(file_name):
    """
    :param file_name: xmind文件名
    :return: xmind转化成的dict数据
    """
    return xmind_to_dict(file_name)[0]


def attach_cases(_original_data):
    """
    :param _original_data: xmind转化成的字典
    :return:按照标题、步骤、预期拼接的case列表
    """
    topic = original_data.get("topic")
    get_xmind_path(topic, '', '')
    cases = []
    for i in range(len(path_list)):
        case_steps = path_list[i].split("#")
        notes = note_list[i].split("#")
        note = get_note(notes)
        expected_result = case_steps[-1]
        title_length = len(get_title(case_steps[2:]))
        if title_length <= 40:
            # title = case_steps[0] + "[%s]" % case_steps[1] + get_title(case_steps[2:])
            title = case_steps[0] + get_title(case_steps[2:])
        else:
            # title = case_steps[0] + "[%s]" % case_steps[1] + get_title(case_steps[2:-1])
            title = case_steps[0] + get_title(case_steps[2:-1])
        step = get_step(case_steps[1:-1])
        case = [title, note, step, expected_result]
        cases.append(case)
    return cases


def get_xmind_path(topic, path, note):
    """
    :param topic: 当前节点后续分支的数据
    :param note: 当前节点添加的备注
    :param path: 当前节点的描述
    :return: 每个分支拼成一个string，返回所有分支的string组成一个list
    """
    path = path + "#" + topic.get("title")
    if topic.get("note", ''):
        # 获取节点上的备注
        note = note + '#' + topic.get("note", '')
    topics = topic.get("topics")
    if not topics:
        # 获取节点的优先级，并且拼在case的最前面
        mark = mark_map.get(topic.get("makers")[0]) if topic.get("makers") else 'P2'
        path = "[%s]%s" % (mark, path)
        path_list.append(path)
        note_list.append(note)
    else:
        for sub_topic in topics:
            get_xmind_path(sub_topic, path, note)



def get_title(title):
    """
    :param title: title部分的拆分list
    :return: 拼接好的string
    """
    length = len(title)
    result = ""
    for each in title:
        result = result + each
        length = length - 1
        if length > 0:
            result += "_"
    return result


def get_step(steps):
    """
    :param steps: step部分的拆分list
    :return: 拼接好的string
    """
    result = ""
    for i, s in enumerate(steps):
        step = "%s.%s\n" % (i+1, s)
        result = result + step
    return result


def get_note(notes):
    """
    :param notes: note部分的拆分list
    :return: 拼接好的string
    """
    result = ""
    notes.remove('')
    if len(notes) == 0:
        return result
    for i, s in enumerate(notes):
        step = "%s.%s\n" % (i+1, s)
        result = result + step
    return result


def write_excel(data):
    timestamp = time.strftime("%Y%m%d%H%M")
    name = "demo" + timestamp + '.xls'
    workbook = xlwt.Workbook(encoding='utf-8')
    data_sheet = workbook.add_sheet('demo')
    row0 = [u'标题', u'前置条件', u'步骤', u'预期结果']
    for i in range(4):
        data_sheet.write(0, i, row0[i])
    for j in range(len(data)):
        for i in range(4):
            data_sheet.write(j+1, i, data[j][i])
    workbook.save(name)


if __name__ == "__main__":
    args = sys.argv
    file_name = args[1]
    original_data = get_original_data(file_name)
    cases = attach_cases(original_data)
    try:
        write_excel(cases)
        print("Successfully")
    except Exception as e:
        print(e)
