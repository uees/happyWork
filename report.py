#!/usr/bin/env python
# -*- coding: utf-8 -*-

import argparse
import os
import random
import re
import sys
from datetime import datetime

from openpyxl import load_workbook

import data_warp as db
import iqc_report
# from fqc_list import FqcListGenerator
from common import is_number, is_number_like, module_path, null2str, rlinput
from config import CONF
from database import Product, init_database, reset_table
from library import WTemplate


class Generator(object):
    ''' 检验报告生成器 '''

    def __init__(self, start_index=None):
        self.start_index = start_index
        self.app_path = module_path()
        self._product_wb_file = "%s/reports/list.xlsx" % self.app_path
        self._wb = load_workbook(self._product_wb_file)
        self._ws = None
        # self.fqc_g = FqcListGenerator('%s/reports/FQC&IQC检测记录.xlsx' % self.app_path)

    def get_start_row(self, ws, flag_done_col=10):
        max_row = ws.max_row
        for row in range(max_row, 2, -1):
            done_cell = ws.cell(row=row, column=flag_done_col)
            if done_cell.value:
                return row + 1

        # if all None, the done row is 2
        return 2

    def generate_reports(self, sheet="Sheet1"):
        ''' 批量生成检验报告 '''
        self._ws = self._wb.get_sheet_by_name(sheet)

        if not self.start_index:
            self.start_index = self.get_start_row(self._ws)

        for index in range(self.start_index, self._ws.max_row + 1):
            print("\n------------------------------")
            self.index = index
            _info = self.get_product_info(index)
            if not _info['internal_name']:
                continue

            print("第 %s行, 品名:%s, 批号:%s" % (index, _info['internal_name'], _info['batch']))

            product = self.query_info(_info)
            if not product:
                continue

            self.generate_report(product)

            # 专用报告
            self.generate_明阳(product)
            self.generate_达进(product)
            self.generate_景旺(product)

            self._set_report_info(product)
            # self.fqc_g.fqc_record(product)

        try:
            # 最後save提高效率 保存采集记录
            self._wb.save(self._product_wb_file)
            # 保存 fqc 记录
            # self.fqc_g.save()
        except PermissionError:
            print("war:文件已经被打开，无法写入")

    def generate_明阳(self, product):
        new_product = {}
        if product['internal_name'].find('A-9060C01') >= 0:
            new_product.update(product)
            new_product["kind"] = "a9060c_my"
            new_product['template'] = self.get_template_by_slug(new_product["kind"])
            new_product["ext_info"] = "(明阳专用报告 粘度为秒)"
            self.generate_report(new_product)

    def generate_达进(self, product):
        new_product = {}
        if product['market_name'] == '8BL2' or product['market_name'] == '8WL5 01':
            new_product.update(product)
            new_product["kind"] = "h8100_dj"
            new_product['template'] = self.get_template_by_slug(new_product["kind"])
            new_product["ext_info"] = "(达进专用报告)"
            self.generate_report(new_product)
        elif product['market_name'] == '44G' or product['market_name'] == '6GHB HF':
            new_product.update(product)
            new_product["kind"] = "h9100_dj"
            new_product['template'] = self.get_template_by_slug(new_product["kind"])
            new_product["ext_info"] = "(达进专用报告)"
            self.generate_report(new_product)

    def generate_景旺(self, product):
        new_product = {}
        if product['market_name'] == '6GHB HF' or product['market_name'] == 'MG55':
            new_product.update(product)
            new_product["kind"] = "h9100_jw"
            new_product['template'] = self.get_template_by_slug(new_product["kind"])
            new_product["ext_info"] = "(景旺专用报告)"
            self.generate_report(new_product)

    def generate_report(self, product):
        ''' 生成检验报告 '''
        template = self.get_template_by_slug(product["kind"])
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

    def get_product_info(self, index):
        ''' 从 list.xlsx 获取指定行的产品输入信息 '''
        validity_date = ''
        ext_info = ''

        customer, internal_name, spec, batch, amount, product_date = [
            self._ws.cell(row=index, column=i).value for i in range(1, 7)]

        customer, internal_name, spec, batch, amount, product_date = map(
            null2str, [customer, internal_name, spec, batch, amount, product_date])

        internal_name = re.sub(r'[\(\)（）]|20kg|20KG|5kg|5KG|1kg|1KG', ' ', internal_name)  # 去除不良字符
        internal_name = internal_name.strip()

        if is_number(batch):
            batch = str(int(batch))
        else:
            print("{}警告：Line {} 批号可能不是数字.{}".format(bcolors.WARNING, index, bcolors.ENDC))

        if len(batch) != 8 and len(batch) != 6:
            ext_info += "(批号格式可能不正确)"
            print("{}Line{}:卧槽,批号格式不一般,小心地雷!!@#$@#&%{}".format(
                bcolors.FAIL, index, bcolors.ENDC))

        if is_number(amount):
            # amount = '发货数量:{}kg'.format(int(amount))
            amount = ''
        else:
            amount = ''

        if not isinstance(product_date, datetime):
            print("{}Line{}:时间格式不正确,已经为你设置为空串.{}".format(
                bcolors.WARNING, index, bcolors.ENDC))
            product_date = ''
            ext_info += "【注意：生产日期没填】"
        else:
            mon_days = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
            next_year = product_date.year
            next_month = product_date.month + 6
            if next_month > 12:
                next_month -= 12
                next_year = product_date.year + 1

            next_day = product_date.day
            if next_day > mon_days[next_month]:
                next_day = mon_days[next_month]

            validity_date = datetime(year=next_year, month=next_month, day=next_day)

            product_date = datetime.strftime(product_date, '%Y/%m/%d')  # 转为字符串
            validity_date = datetime.strftime(validity_date, '%Y/%m/%d')  # 转为字符串

        return dict(customer=customer,
                    internal_name=internal_name,
                    spec=spec,
                    batch=batch,
                    amount=amount,
                    product_date=product_date,
                    validity_date=validity_date,
                    ext_info=ext_info)

    def _input_kind(self):
        kind = rlinput("类别(H-8100/H-9100/A-2000/K-2500/A-2100/A-9060A/A-9000/\nUVS-1000/TM-3100/TS-3000/UVM-1800):\n >>>")

        conf = self.get_conf(kind)
        if not conf:
            conf = self.get_conf_by_alias(kind)
        if not conf:
            print("无效的类别, 请重新输入")
            conf = self._input_kind()

        return conf

    def _input_number(self, msg):
        value = rlinput(msg)

        if not is_number_like(value):
            value = self._input_number()

        return value

    def _cmd_add(self, internal_name):
        market_name = rlinput("销售品名(打在检验报告上的名字，例如‘8G04建业’的销售名为8G 04):\n >>>")
        conf = self._input_kind()
        viscosity = self._input_number("粘度值")
        viscosity_width = self._input_number("粘度上下幅度:")

        product_obj = Product(internal_name=internal_name,
                              market_name=market_name,
                              template=conf['slug'],
                              viscosity=int(viscosity),
                              viscosity_width=int(viscosity_width))
        db.insert_product(product_obj)
        db.insert_product_to_xlsx(product_obj, '%s/data/database.xlsx' % self.app_path, 'products')
        print("已经插入新的条目到products数据表")
        return product_obj

    def _cmd_edit(self, given):
        given['internal_name'] = rlinput("品名:", given['internal_name'])
        self._ws.cell('B{}'.format(self.index)).value = given['internal_name']
        try:
            self._wb.save(self._product_wb_file)
        except PermissionError:
            print("war:文件已经被打开，无法写入")

    def _validate_id(self, id, list_ids):
        ''' 验证指定的ID是否在给定的列表中 '''
        if not re.match(r"\d+$", id):
            print("输入的ID不是一个数字")
            return False
        if int(id) not in list_ids:
            print("输入的ID不在指定的范围")
            return False
        return True

    def query_info(self, given):
        '''
        查询数据库, 追加更多信息
        :param: given 是从list.xlsx 中查出的数据
        '''
        products = db.search_product(given['internal_name'])
        if products.count() == 0:
            print('数据库中无记录,请输入命令. \n 添加条目：add \n 跳过此行：break \n 编辑字段：edit \n 退出：任意其他字符')
            cmd = rlinput("Command:")
            if cmd == "add":
                product_obj = self._cmd_add(given['internal_name'])

            elif cmd == "break":
                return

            elif cmd == "edit":
                self._cmd_edit(given)
                return self.query_info(given)

            else:
                sys.exit()

        elif products.count() == 1:
            product_obj = products.one()

        else:
            ids = []
            for product in products.all():
                ids.append(product.id)
                space = " " * (20 - len(product.internal_name)) if len(product.internal_name) < 20 else ""
                print("\t %s%s\t ID:%s\t %s±%sdPa.s" % (product.internal_name,
                                                        space,
                                                        product.id,
                                                        product.viscosity,
                                                        product.viscosity_width))

            while True:
                print("小提示: 你可以输入 qui t立即退出, 要编辑字段输入 edit")
                pid = rlinput("please choise a ID:")
                if pid == "quit":
                    sys.exit()

                elif pid == "edit":
                    self._cmd_edit(given)
                    return self.query_info(given)

                if self._validate_id(pid, ids):
                    break

            product_obj = db.get_product_by_id(pid)
            print("\t 你选择了(%s, %s±%sdPa.s)" % (product_obj.internal_name,
                                               product_obj.viscosity,
                                               product_obj.viscosity_width))

        given['market_name'] = product_obj.market_name
        given['kind'] = product_obj.template
        given['template'] = self.get_template_by_slug(given["kind"])
        given['viscosity'] = product_obj.viscosity
        given['viscosity_limit'] = "%s±%s" % (product_obj.viscosity, product_obj.viscosity_width)
        given['qc_date'] = datetime.strftime(datetime.now(), '%Y/%m/%d')
        given['ftir'] = '{}%'.format(round(random.uniform(99.3, 100), 2))

        given = self.given_修饰(given, product_obj)

        return given

    def given_修饰(self, given, product_obj):

        # 征对性修饰
        if product_obj.market_name.find('SP8') == 0 or \
                product_obj.market_name == 'A-9060A 01' or \
                product_obj.market_name == '60G':
            given['ext_info'] += '(深南电路要求打发货数量)'

        if product_obj.market_name.find('28GHB') >= 0 or \
                product_obj.market_name.find('30GHB') >= 0 or \
                product_obj.market_name.find('SP20HF') >= 0:
            given['ext_info'] += '(宏华胜要求打发货数量)'

        if product_obj.market_name == '8BL' or \
                product_obj.market_name == 'GH3' or \
                product_obj.market_name.find('G6') == 0 or \
                product_obj.market_name.find('SP02') == 0 or \
                product_obj.market_name.find('GH40') == 0 or \
                product_obj.market_name.find('MG31') == 0 or \
                product_obj.market_name.find('23GHB') == 0 or \
                product_obj.internal_name.find('崇达') >= 0:
            given['ext_info'] += '(崇达要求打发货数量)'

        return given

    def _set_report_info(self, product):
        ''' 写入部分信息到指定行 '''
        self._ws.cell('G{}'.format(
            self.index)).value = product['internal_name']
        self._ws.cell('H{}'.format(
            self.index)).value = product['viscosity_limit']
        self._ws.cell('I{}'.format(
            self.index)).value = product['product_date']
        self._ws.cell('J{}'.format(
            self.index)).value = product['validity_date']
        self._ws.cell('K{}'.format(
            self.index)).value = product['qc_date']

    def get_today_report_dir_path(self):
        '''自动创建并返回当日报告文件夹路径'''
        today_dir_name = datetime.strftime(datetime.now(), '%Y-%m-%d')
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

    def get_template_by_slug(self, slug):
        conf = self.get_conf(slug)
        return self.get_template(conf.get('template'))

    def get_conf(self, slug):
        for item in CONF:
            if item.get('slug') == slug:
                return item

    def get_conf_by_alias(self, alias):
        for item in CONF:
            if alias in item.get('alias'):
                return item

    def get_template(self, name):
        ''' 获取模板文件路径 '''
        return os.path.join(self.app_path, "templates/%s.docx" % name)


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
    parser.add_argument("--create_all", action="store_true",
                        default=False, help="初始化数据库")
    parser.add_argument("--reset_table", help="重置单个表")
    parser.add_argument("--init_products", action="store_true",
                        default=False, help="从xlsx文件中采集数据")
    parser.add_argument("--init_materials", action="store_true",
                        default=False, help="从xlsx文件中采集IQC检测要求数据")

    group = parser.add_argument_group('iqc')
    group.add_argument("--iqc", action="store_true",
                       default=False, help="创建IQC报告")
    group.add_argument('-f', "--filename", help="IQC流水文件")
    group.add_argument('-e', "--end_row", type=int, help="IQC流水文件结束行")
    # subparsers = parser.add_subparsers(help='commands')
    # iqc_parser = subparsers.add_parser('iqc', help='创建IQC报告')
    # iqc_parser.add_argument('-f', "--filename", help="IQC流水文件")
    # iqc_parser.add_argument('-e', "--end_row", type=int, help="IQC流水文件结束行")

    args = parser.parse_args()

    if args.index:
        g = Generator(args.index)
        g.generate_reports()

    elif args.create_all:
        init_database()

    elif args.reset_table:
        reset_table(args.reset_table)

    elif args.init_products:
        reset_table(Product)
        db.init_product_data('{}/data/database.xlsx'.format(module_path()), 'products')

    elif args.init_materials:
        reset_table('iqc_materials')
        db.init_materials('{}/data/原材料检验项目及要求.xlsx'.format(module_path()), 'Sheet1')

    elif args.iqc:
        if args.filename:
            iqc_report.generate(args.filename, args.end_row)
        else:
            print("filename parameter must be provided. please refer --help for more detail.")
    else:
        g = Generator()
        g.generate_reports()
