#!/usr/bin/env python
#-- coding: utf-8 -*-

import os
import re
import sys
import argparse
from datetime import datetime
from openpyxl import load_workbook
from config import TEMPLATES
from common import module_path, is_number
from database import Product
from library import WTemplate
import data_warp as db


class Generator(object):
    ''' 检验报告生成器 '''
    def __init__(self, start_index):
        self.app_path = module_path()
        self.start_index = start_index
        self._product_wb_file = "%s/reports/list.xlsx" % self.app_path
        self._wb = load_workbook(self._product_wb_file)
        self._ws = None
        self._product = dict()
        
        
    def generate_report(self, product):
        ''' 生成检验报告 '''
        template = self.get_template(product["template"])
        if not os.path.exists(template):
            print("%s 模板文件不存在！" % product["kind"])
            return
        tp = WTemplate(template)
        tp.replace(product)

        today_report_dir = self.get_today_report_dir_path()
        filename = '{}_{}_{}{}.docx'.format(product["batch"],
                                            product["internal_name"],
                                            product["spec"],
                                            product["ext_info"])
        filepath = '{}/{}'.format(today_report_dir, filename)
        
        if os.path.exists(filepath):
            print("{}{}已经存在了{}".format(bcolors.WARNING, filename, bcolors.ENDC))
        else:
            tp.save(filepath)
            print("报告已经生成：{}".format(filename))


    def generate_reports(self, sheet="Sheet1"):
        ''' 批量生成检验报告 '''
        self._ws = self._wb.get_sheet_by_name(sheet)
        for index in range(self.start_index, self._ws.max_row+1):
            print("\n------------------------------")
            _info = self._get_product_info(index)
            if not _info['internal_name']:
                continue
            print("第 %s行, 品名:%s, 批号:%s" % (index, _info['internal_name'], _info['batch']))
            
            product = self._query_info(_info)
            
            self.generate_report(product)
            
            self._set_report_info(index, product)

    def _validate_id(self, id, list_ids):
        ''' 验证指定的ID是否在给定的列表中 '''
        if not re.match(r"\d+$", id):
            print("输入的ID不是一个数字")
            return False
        if int(id) not in list_ids:
            print("输入的ID不在指定的范围")
            return False
        return True
    
    def _get_product_info(self, index):
        ''' 获取指定行的产品输入信息 '''
        customer, \
        internal_name, \
        spec, \
        batch, \
        amount, \
        product_date = [self._ws.cell(row=index, column=i).value for i in range(1, 7)]
        
        validity_date = ""
        ext_info = ''
            
        if customer is None:
            customer = ""
        
        pattern = re.compile(r'[\(\)（）]|20kg|20KG|5kg|5KG|1kg|1KG')
        internal_name, number = pattern.subn(' ', internal_name)  #去除不良字符
        internal_name = internal_name.strip()
        
        if spec is None:
            spec = ""

        if batch is None:
            batch = ""
        elif is_number(batch):
            batch = str(int(batch))
            if len(batch)!=8 and len(batch)!=6:
                ext_info += "(批号不是6位或8位数字)"
                print("{}Line{}:卧槽,批号格式不一般,小心地雷!!@@{}".format(bcolors.WARNING, index, bcolors.ENDC))
        else:
            ext_info += "(批号可能不是数字)"
            print("{}警告：Line {} 批号可能不是数字.{}".format(bcolors.WARNING, index, bcolors.ENDC))
        
        if is_number(amount):
            #amount = '发货数量:{}kg'.format(int(amount))
            amount = ''
        else:
            amount = ''
            
        if not isinstance(product_date, datetime):
            print("{}Line{}:时间格式不正确,已经为你设置为空串.{}".format(bcolors.WARNING, index, bcolors.ENDC))
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
            
            product_date = datetime.strftime(product_date,'%Y-%m-%d')  #转为字符串
            validity_date = datetime.strftime(validity_date,'%Y-%m-%d')  #转为字符串
            
        return dict(customer = customer,
                    internal_name = internal_name,
                    spec = spec,
                    batch = batch,
                    amount = amount,
                    product_date = product_date,
                    validity_date = validity_date,
                    ext_info = ext_info)
    
               
    def _query_info(self, given):
        ''' 查询数据库,追加更多信息 '''
        products = db.search_product(given['internal_name'])
        if products.count() == 1:
            product_obj = products.first()
        elif products.count() == 0:
            print("数据库中无记录, 添加条目请按输入add, 若要退出请输入任意字符")
            cmd = input("Command:")
            if cmd == "add":
                market_name = input("销售品名:")
                template = input("模板:")
                viscosity = input("粘度:")
                viscosity_width = input("粘度幅度:")
                product_obj = Product(internal_name=given['internal_name'],
                                      market_name=market_name,
                                      template=template,
                                      viscosity=int(viscosity),
                                      viscosity_width=int(viscosity_width))
                db.insert_product(product_obj)
                print("已经插入新的条目到products数据表")
                db.insert_product_to_xlsx(product_obj, 
                                          "%s/data/database.xlsx" % self.app_path, 
                                          'Sheet1')
            else:
                sys.exit()
        else:
            list_ids = list()
            for product in products.all():
                list_ids.append(product.id)
                space = " " * (20-len(product.internal_name)) if len(product.internal_name)<20 else ""
                print("\t %s%s\t ID:%s\t %s±%sdPa.s" % (product.internal_name, 
                                                        space, 
                                                        product.id, 
                                                        product.viscosity, 
                                                        product.viscosity_width))
            
            while True:
                print("小提示: 你可以输入quit立即退出")
                pid = input("please choise a ID:")
                if pid == "quit":
                    sys.exit()
                if self._validate_id(pid, list_ids):
                    break
                
            product_obj = db.get_product_by_id(pid)
            print("\t 你选择了(%s, %s±%sdPa.s)" % (product_obj.internal_name, 
                                                product_obj.viscosity, 
                                                product_obj.viscosity_width)
                  )
        
        if product_obj.market_name == 'SP8' or product_obj.market_name == 'A-9060A 01':
            given['ext_info'] += '(深南电路要求打发货数量)'
        
        given['market_name'] = product_obj.market_name
        given['template'] =TEMPLATES[product_obj.template]
        given['viscosity'] =product_obj.viscosity
        given['viscosity_limit'] = "%s±%s" % (product_obj.viscosity, product_obj.viscosity_width)
        given['qc_date'] = datetime.strftime(datetime.now(),'%Y-%m-%d')
        return given
    
    
    def _set_report_info(self, index, product):
        ''' 写入部分信息到指定行 '''
        self._ws.cell('G{}'.format(index)).value = product['internal_name']
        self._ws.cell('H{}'.format(index)).value = product['viscosity_limit']
        self._ws.cell('I{}'.format(index)).value = product['product_date']
        self._ws.cell('J{}'.format(index)).value = product['validity_date']
        self._ws.cell('K{}'.format(index)).value = product['qc_date']
        self._wb.save(self._product_wb_file)
    
    
    def get_today_report_dir_path(self):
        '''自动创建并返回当日报告文件夹路径'''
        today_dir_name = datetime.strftime(datetime.now(),'%Y-%m-%d')
        today_path = os.path.join(module_path(), 'reports/%s' % today_dir_name)
        if not os.path.exists(today_path):
            os.mkdir(today_path)
        return today_path
    
    
    def onlyone_filename(self, path, filename, ext):
        '''检查并生成唯一的文件名'''
        filepath = '%s/%s.%s' % (path, filename, ext)
        if os.path.exists(filepath):
            filename = '%s(1)' % filename
            filename = self.onlyone_filename(path, filename, ext)
        return filename
    
    
    def get_template(self, name):
        ''' 获取模板文件路径 '''
        template = os.path.join(self.app_path, "templates/%s.docx" % name)
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
    parser.add_argument("--init_database", action="store_true", default=False, help="初始化数据库")
    parser.add_argument("--reset_table", help="重置单个表")
    parser.add_argument("--fetch_product_data", action="store_true", 
                        default=False, help="从xlsx文件中采集数据")
    args = parser.parse_args()
    if args.index:
        g = Generator(args.index)
        g.generate_reports()
    elif args.init_database:
        db.init_database()
    elif args.reset_table:
        db.reset_table(args.reset_table)
    elif args.fetch_product_data:
        db.reset_table(Product)
        db.fetch_product_data('{}/data/database.xlsx'.format(module_path()), 'products')
    else:
        print("unknown parameter, please refer --help for more detail.")