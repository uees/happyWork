#!/usr/bin/env python
#-- coding: utf-8 -*-

import os
import re
import sys
import argparse
import time
from datetime import datetime
from officer import Word, Excel
from config import TEMPLATES, REPORT_PATH
from common import exe_path, Objdict, is_number

class Generator(object):
    def __init__(self, start, file_format):
        self.start_index = start
        self.file_format = file_format
        self.app_path = exe_path()
        if REPORT_PATH:
            if not os.path.exists("%s\\list.xlsx" % REPORT_PATH):
                print("%s\\list.xlsx 不存在" % REPORT_PATH)
                sys.exit()
            self.excel = Excel("%s\\list.xlsx" % REPORT_PATH)
        else:
            self.excel = Excel("%s\\reports\\list.xlsx" % self.app_path)
            
        self.tm_year = time.localtime().tm_year  #int
        self.tm_mon = time.localtime().tm_mon  #int
        
        
    def generate_report(self, product):
        template = get_template(product["template"])
        if not template:
            print("%s 模板文件不存在！" % product["kind"])
            return
        docer = Word(template)
        docer.replace(product)

        today_report_dir = get_today_report_dir_path(REPORT_PATH)
        filename = '{}_{}_{}{}.{}'.format(product["batch"],
                                           product["raw_name"],
                                           product["spec"],
                                           product["ext_info"],
                                           self.file_format)
        filepath = r'{}\{}'.format(today_report_dir, filename)
        
        if os.path.exists(filepath):
            print("{}{}已经存在了{}".format(bcolors.WARNING, filename, bcolors.ENDC))
            docer.close(0)
        else:
            if self.file_format == 'pdf':
                docer.ExportAsPDF(filepath)
            else:
                docer.SaveAs(filepath)
            docer.close(0)
            print("报告已经生成：{}".format(filename))


    def generate_reports(self):
        sheet = "Sheet1"
        for i in range(self.start_index, self.excel.get_numrows(sheet)+1):
            print("\n------------------------------")
            
            customer, name, spec, batch, amount, product_date =\
                self.excel.getRange(sheet, i, 1, i, 6)[0]
            validity_date = ""
            today = datetime.strftime(datetime.now(),'%Y-%m-%d')
            ext_info = ''
                
            if customer is None:
                customer = ""
                
            if spec is None:
                spec = ""
                
            if not name: 
                continue
            
            pattern = re.compile(r'[\(\)（）]|20kg|20KG|5kg|5KG|1kg|1KG')
            name, number = pattern.subn(' ', name)  #去除不良字符
            name = name.strip()
            

            if batch is None:
                batch = ""
            elif is_number(batch):
                batch = str(int(batch))
                if len(batch)!=8 and len(batch)!=6:
                    ext_info += "(批号不是6位或8位数字)"
                    print("{}Line{}:卧槽,批号格式不一般,小心地雷!!@@{}".format(bcolors.WARNING, i, bcolors.ENDC))
            else:
                ext_info += "(批号可能不是数字)"
                print("{}警告：Line {} 批号可能不是数字.{}".format(bcolors.WARNING, i, bcolors.ENDC))
            
            if is_number(amount):
                #amount = '发货数量:{}kg'.format(int(amount))
                amount = ''
            else:
                amount = ''
                
            if not isinstance(product_date, datetime):
                print("{}Line{}:时间格式不正确,已经为你设置为空串.{}".format(bcolors.WARNING, i, bcolors.ENDC))
                product_date = ""
                ext_info += "【注意：生产日期没填】"
            else:
                mon_days = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
                next_year = product_date.year
                next_month = product_date.month + 6
                if next_month > 12:
                    next_month -= 12
                    next_year = product_date.year +1
                    
                next_day = product_date.day
                if next_day > mon_days[next_month]:
                    next_day = mon_days[next_month]
                    
                validity_date = datetime(year=next_year, month=next_month, day=next_day)
                
                product_date = datetime.strftime(product_date,'%Y-%m-%d')
                validity_date = datetime.strftime(validity_date,'%Y-%m-%d')
                
                 
            print("第 %s 行, 客户:%s, 品名:%s, 批号:%s" % (i, customer, name, batch))
            products = db.search(name)
            if products.count == 1:
                product_obj = products.index(0)
            elif products.count == 0:
                print("数据库中无记录, 添加条目请按输入add, 若要退出请输入任意字符")
                cmd = input("Command:")
                if cmd == "add":
                    slug = input("Slug:")
                    kind = input("Categry:")
                    niandu = input("niandu:")
                    limit = input("limit:")
                    product_obj = Objdict(name=name,
                                          slug=slug,
                                          kind=kind,
                                          niandu=int(niandu),
                                          limit=int(limit))
                    db.insert_product(product_obj)
                    
                    db_excel = Excel("%s\\data\\db.xlsx" % self.app_path)
                    index = db_excel.get_numrows("Sheet3")+1
                    db_excel.setCell("Sheet3", index, 1, name)
                    db_excel.setCell("Sheet3", index, 2, kind)
                    db_excel.setCell("Sheet3", index, 5, niandu)
                    db_excel.setCell("Sheet3", index, 6, limit)
                    db_excel.setCell("Sheet3", index, 7, slug)
                    db_excel.close()
                    print("已经插入新的条目到products数据表")
                else:
                    sys.exit()
            else:
                list_ids = list()
                for product in products:
                    list_ids.append(product.id)
                    space = " " * (20-len(product.name)) if len(product.name)<20 else ""
                    print("\t %s%s\t ID:%s\t %s±%sdPa.s" % (product.name, 
                                                            space, 
                                                            product.id, 
                                                            product.niandu, 
                                                            product.limit)
                          )
                
                while True:
                    print("小提示: 你可以输入quit立即退出")
                    pid = input("please choise a ID:")
                    if pid == "quit":
                        sys.exit()
                    if self._validate_id(pid, list_ids):
                        break
                    
                product_obj = db.get_product_by_id(pid)
                print("\t 你选择了(%s, %s±%sdPa.s)" % (product_obj.name, 
                                                    product_obj.niandu, 
                                                    product_obj.limit)
                      )
            
            if product_obj.slug == 'SP8' or product_obj.slug == 'A-9060A 01':
                ext_info += '(深南电路要求打发货数量)'
                
            product = dict(raw_name=name,
                           name=product_obj.slug,
                           spec=spec,
                           template=TEMPLATES[product_obj.kind],
                           batch=batch, 
                           niandu=product_obj.niandu,
                           niandu_limit = "%s±%s" % (product_obj.niandu, product_obj.limit),
                           amount = amount,
                           customer = customer,
                           product_date = product_date,
                           validity_date = validity_date,
                           qc_date = today,
                           ext_info=ext_info)
            self.generate_report(product)
            
            self.excel.setRange(sheet, 7, i, [(product_obj.name, 
                                               product["niandu_limit"], 
                                               product_date,
                                               validity_date,
                                               today)])
            self.excel.save()
            
            # 插入批次信息
            #db.insert_batch(Objdict(name=product_obj.name, batch=batch, qc_date=today))

    def _validate_id(self, id, list_ids):
        if not re.match(r"\d+$", id):
            print("输入的ID不是一个数字")
            return False
        if int(id) not in list_ids:
            print("输入的ID不在指定的范围")
            return False
        return True
    
    
def get_today_report_dir_path(base_dir=""):
    '''自动创建并返回当日报告文件夹路径'''
    today_dir_name = datetime.strftime(datetime.now(),'%Y-%m-%d')
    if base_dir:
        today_path = '%s\\%s' % (base_dir, today_dir_name)
    else:
        today_path = os.path.join(exe_path(), 'reports\\%s' % today_dir_name)
    if not os.path.exists(today_path):
        os.mkdir(today_path)
    return today_path


def onlyone_filename(path, filename, ext):
    '''检查并生成唯一的文件名'''
    filepath = '%s\\%s.%s' % (path, filename, ext)
    if os.path.exists(filepath):
        filename = '%s(1)' % filename
        filename = onlyone_filename(path, filename, ext)
    return filename


def get_template(name):
    template = os.path.join(exe_path(), "templates\\%s.doc" % name)
    if not os.path.exists(template):
        return None
    return template

class bcolors:
    HEADER = '\033[95m'           # 粉色
    OKBLUE = '\033[94m'           # 蓝色
    OKGREEN = '\033[92m'          # 绿色
    WARNING = '\033[93m'          # 黄色
    FAIL = '\033[91m'             # 红色
    BOLD = '\033[1m'              # 粗体
    ENDC = '\033[0m'              # 结束


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--index", type=int, help="excel中需要生成报告的起始行")
    parser.add_argument("-f", "--format", choices=['pdf', 'doc'], default='doc', help="格式化为PDF还是DOC")
    parser.add_argument("--init_database", action="store_true", default=False, help="初始化数据库")
    parser.add_argument("--reset_table", help="重置单个表")
    args = parser.parse_args()
    if args.index:
        g = Generator(args.index, args.format)
        g.generate_reports()
    elif args.init_database:
        init_database()
    elif args.reset_table:
        reset_table(get_table_class(args.reset_table))
    else:
        print("unknown parameter, please refer --help for more detail.")