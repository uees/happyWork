# -*- coding:utf-8 -*-
'''
Created on 2016年5月28日

@author: QC
'''

import os
import shutil

from common import module_path

app_path = module_path()
report_path = os.path.join(app_path, 'reports')
output_dir = os.path.join(app_path, '搜索结果')


def get_all_files(dir_path):
    ''' 递归获取文件夹下所有的文件 '''
    files = list()
    for item in os.listdir(dir_path):
        item_path = os.path.join(dir_path, item)
        if os.path.isdir(item_path):
            files.extend(get_all_files(item_path))
        else:
            files.append(item_path)
    return files


def search_and_copy(batch, files, finded_break=False):
    ''' 按批号搜索并拷贝 '''
    find = False
    for file_path in files:
        if file_path.find(batch) >= 0:
            find = True
            try:
                shutil.copy(file_path, output_dir)
            except shutil.SameFileError:
                pass
            if finded_break:
                break
    if not find:
        print('批号{}的报告没找到'.format(batch))
        with open(os.path.join(output_dir, 'log.txt'), 'a') as fp:
            fp.write('批号{}的报告没找到\n'.format(batch))


def run():
    ''' 读取'批号.txt'中的文本执行查找 '''
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)
    docs = get_all_files(report_path)
    with open(os.path.join(app_path, '批号.txt'), 'r') as fp:
        for batch in fp.readlines():
            batch = batch.strip()
            if len(batch) == 0:  # 判断是否是空行
                continue
            search_and_copy(batch, docs)


if __name__ == '__main__':
    print("开始搜索。。。")
    run()
    print("搜索完毕!")
