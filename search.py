#-*- coding:utf-8 -*-
'''
Created on 2016年5月28日

@author: QC
'''

import glob
import os
import shutil
from common import module_path

app_path = module_path()
report_path = '{}/reports'.format(app_path)
output_dir = '{}/搜索结果'.format(app_path)

def get_all_docs():
    docs = glob.glob("{}/*.doc*".format(report_path))  # a list  .doc  .docx
    for dir_name in os.listdir(report_path):
        dir_path = "{}/{}".format(report_path, dir_name)
        if os.path.isdir(dir_path): 
            docs.extend(glob.glob("{}/*.doc*".format(dir_path)))
    return docs
            
def search(batch, docs):
    is_find = False
    for path_and_filename in docs:
        if path_and_filename.find(batch) >= 0:
            is_find = True
            try:
                shutil.copy(path_and_filename, output_dir)
            except shutil.SameFileError:
                pass
    if not is_find:
        print('批号{}的报告没找到'.format(batch))
        with open('{}/log.txt'.format(output_dir), 'a') as fp:
            fp.write('批号{}的报告没找到\n'.format(batch))
    
def run():
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)
    docs = get_all_docs()
    with open('{}/批号.txt'.format(app_path), 'r') as fp:
        for batch in fp.readlines():
            batch = batch.strip()
            if not len(batch): #判断是否是空行
                continue
            search(batch, docs)
    print("搜索完毕!")
    
        
if __name__ == '__main__':
    run()