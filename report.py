#!/usr/bin/env python
# -- coding: utf-8 -*-

import os
import re
import sys
import argparse
import random
import data_warp as db
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font
from common import module_path, is_number, rlinput, null2str
from database import Product, IQCMaterial, init_database, reset_table
from library import WTemplate
from config import TEMPLATES, ALL_FQC_ITEMS, FQC_ITEMS


def generate_iqc_reports(filename, end_row=None):
    wb = load_workbook(filename)
    ws = wb.get_sheet_by_name('供应商来料质量统计表')
    if not end_row:
        end_row = ws.max_row
    for row in ws.iter_rows('B7:G{}'.format(end_row)):
        template_wb = load_workbook('%s/templates/iqc.xlsx' % module_path())
        template_ws = template_wb.get_sheet_by_name('Sheet1')
        material_name, incoming_date, supplier, qc_result, substandard_items, \
            amount = [cell.value for cell in row]

        if isinstance(material_name, str):
            _material_name = re.sub(
                r'[\u4e00-\u9fa5]+', '', material_name)  # 去掉中文
            if not _material_name:
                continue
        material = IQCMaterial.query.filter(
            IQCMaterial.name.ilike('%' + _material_name + '%')).first()
        if not material or material.qc_items == '免检':
            continue

        if isinstance(incoming_date, datetime):
            incoming_date = datetime.strftime(
                incoming_date, '%Y-%m-%d')  # 转为字符串

        if _material_name.upper() in ['0.25L', '0.3L', '1L', '5L', '6L', '20L',
                                      '0.25KG', '0.3KG', '1KG', '5KG', '6KG',
                                      '20KG']:
            unit = '套'
        else:
            unit = 'kg'

        template_ws.cell('B5').value = incoming_date
        template_ws.cell('B6').value = material_name
        template_ws.cell('D6').value = supplier
        template_ws.cell('D7').value = '%s%s' % (amount, unit)
        template_ws.cell('D8').value = material.qc_method

        qc_items = material.qc_items.split('、')
        row = 11
        for item in qc_items:
            template_ws.cell('A{}'.format(row)).value = item
            if item == '细度':
                if _material_name.find('A0084') >= 0:
                    template_ws.cell('C{}'.format(row)).value = '<25μm'
                elif _material_name.find('A0085') >= 0 or \
                        _material_name.find('A0088') >= 0:
                    template_ws.cell('C{}'.format(row)).value = '<17.5μm'
                else:
                    template_ws.cell('C{}'.format(row)).value = '<20μm'
            elif item == '软化点':
                if _material_name == 'A0016' or _material_name == 'A0016A'\
                        or _material_name == 'A0016B':
                    template_ws.cell('C{}'.format(row)).value = \
                        '%s℃' % round(random.uniform(27, 30), 1)
                elif _material_name == 'A0016F':
                    template_ws.cell('C{}'.format(row)).value = \
                        '%s℃' % round(random.uniform(32, 35), 1)
                else:
                    template_ws.cell('C{}'.format(row)).value = '√'
            elif item == '环氧值':
                if _material_name == 'A0016' or _material_name == 'A0016A'\
                        or _material_name == 'A0016B':
                    template_ws.cell('C{}'.format(row)).value = \
                        '%s mol/100g' % round(random.uniform(0.515, 0.535), 3)
                elif _material_name == 'A0016F':
                    template_ws.cell('C{}'.format(row)).value = \
                        '%s mol/100g' % round(random.uniform(0.56, 0.59), 3)
                else:
                    template_ws.cell('C{}'.format(row)).value = '√'
            elif item == '馏程':
                if _material_name.find('A0055') >= 0 or \
                        _material_name.find('A0063') >= 0:
                    template_ws.cell('C{}'.format(row)).value = \
                        '%s~%s℃' % (str(180 + random.randint(1, 9)),
                                    str(220 - random.randint(1, 9)))
                elif _material_name.find('A0058') >= 0:
                    template_ws.cell('C{}'.format(row)).value = \
                        '%s~%s℃' % (str(135 + random.randint(1, 5)),
                                    str(150 - random.randint(1, 5)))
                else:
                    template_ws.cell('C{}'.format(row)).value = '√'
            else:
                template_ws.cell('C{}'.format(row)).value = '√'
            row += 1
        template_ws.merge_cells('B11:B{}'.format(10 + len(qc_items)))
        template_ws.cell('B11').value = material.spec

        new_filename = '%s-%s-%s-%s.xlsx' % (incoming_date,
                                             random.randint(1, 99),
                                             material_name,
                                             supplier)
        template_wb.save('%s/reports/IQC/%s' % (module_path(), new_filename))


class Generator(object):
    ''' 检验报告生成器 '''

    def __init__(self, start_index):
        self.app_path = module_path()
        self.start_index = start_index
        self._product_wb_file = "%s/reports/list.xlsx" % self.app_path
        self._wb = load_workbook(self._product_wb_file)
        self._ws = None
        self._font = Font(name='Calibri', size=10)

    def generate_reports(self, sheet="Sheet1"):
        ''' 批量生成检验报告 '''
        self._ws = self._wb.get_sheet_by_name(sheet)
        fqc_filename = '%s/reports/FQC&IQC检测记录.xlsx' % self.app_path
        fqc_wb = load_workbook(fqc_filename)
        fqc_ws = fqc_wb.get_sheet_by_name('FQC')
        row = fqc_ws.max_row
        for index in range(self.start_index, self._ws.max_row + 1):
            self.index = index
            print("\n------------------------------")
            _info = self._get_product_info(index)
            if not _info['internal_name']:
                continue
            print("第 %s行, 品名:%s, 批号:%s" %
                  (index, _info['internal_name'], _info['batch']))

            product = self._query_info(_info)
            if not product:
                continue

            product_dj = dict()
            if product['market_name'] == '8BL2' or \
                    product['market_name'] == '8WL501':
                product_dj.update(product)
                product_dj["kind"] = "h8100_dj"
                product_dj['template'] = TEMPLATES["h8100_dj"]
                product_dj["ext_info"] = "(达进专用报告)"
                self.generate_report(product_dj)
            elif product['market_name'] == '44G' or \
                    product['market_name'] == '6GHB':
                product_dj.update(product)
                product_dj["kind"] = "h9100_dj"
                product_dj['template'] = TEMPLATES["h9100_dj"]
                product_dj["ext_info"] = "(达进专用报告)"
                self.generate_report(product_dj)

            self.generate_report(product)
            self._set_report_info(product)

            row = row + 1
            self.fqc_record(product, fqc_ws, row)

        try:  # 最後save提高效率
            self._wb.save(self._product_wb_file)
            fqc_wb.save(fqc_filename)
        except PermissionError:
            print("war:文件已经被打开，无法写入")

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

    def fqc_record(self, product, ws, row):
        record = self._make_record(product)
        if not record:
            return
        for (i, item) in zip(range(1, 5), [product['product_date'],
                                           product['qc_date'],
                                           product['internal_name'],
                                           product['batch']
                                           ]):
            ws.cell(row=row, column=i).value = item
            ws.cell(row=row, column=i).font = self._font

        for item in record:
            ws.cell(row=row, column=item['col']).value = item['value']
            ws.cell(row=row, column=item['col']).font = self._font

    def _make_record(self, product):
        record = []
        c = self._get_cat(product["kind"])
        if not c:
            return
        items = FQC_ITEMS.get(c)
        for item in items:
            if item == '外观颜色' or item == '板面效果' or item == '固化性'\
                    or item == '显影性' or item == '去膜性'\
                    or item == '耐焊性' or item == '耐化学性':
                record.append({'col': self._get_col(item), 'value': '√'})
            elif item == '细度':
                record.append({'col': self._get_col(item),
                               'value': '%sμm' % random.choice([15, 17.5, 20])})
            elif item == '反白条':
                if c == 'uvw5d10' or c == 'uvw5d65' or c == 'a2' or c == 'k2'\
                        or c == 'thw5d35':
                    record.append(
                        {'col': self._get_col(item), 'value': '≤10μm'})
                elif c == 'tm3100' or c == 'ts3000' or c == 'uvs1000' or c == 'uvm1800':
                    record.append(
                        {'col': self._get_col(item), 'value': '≤20μm'})
                else:
                    record.append(
                        {'col': self._get_col(item), 'value': '≤5μm'})
            elif item == '粘度':
                if product['internal_name'].find('A-9060C明阳') >= 0:
                    unit = 's/32℃'
                else:
                    unit = 'dpa.s/25℃'
                record.append({'col': self._get_col(item),
                               'value': '%s%s' % (self._get_viscosity(product['viscosity']), unit)})
            elif item == '硬度':
                if c == 'uvw5d10':
                    record.append({'col': self._get_col(item), 'value': 'H'})
                elif c == 'a9':
                    record.append({'col': self._get_col(item), 'value': '2H'})
                elif c == 'uvm1800':
                    record.append({'col': self._get_col(item), 'value': '3H'})
                elif c == 'tm3100' or c == 'uvs1000':
                    record.append({'col': self._get_col(item), 'value': '4H'})
                elif c == 'ts3000':
                    record.append({'col': self._get_col(item), 'value': '5H'})
                elif c == 'h9100' or c == 'h8100':
                    record.append({'col': self._get_col(item), 'value': '6H'})
            elif item == '附着力':
                record.append({'col': self._get_col(item), 'value': '100%'})
            elif item == '感光性':
                if c == 'h9100' or c == 'h8100':
                    record.append({'col': self._get_col(item),
                                   'value': '%s级' % random.choice([9.5, 10, 10.5, 11])})
                elif c == 'a9':
                    record.append({'col': self._get_col(item),
                                   'value': '%s级' % random.choice([6, 6.5, 7])})
            elif item == '解像性':
                record.append({'col': self._get_col(item), 'value': '≤50μm'})
            elif item == '红外图谱':
                record.append({'col': self._get_col(item),
                               'value': '{}%'.format(round(random.uniform(99.3, 100), 2))})
        return record

    def _get_viscosity(self, viscosity):
        viscosity = int(viscosity)
        if viscosity <= 30:
            viscosity += random.uniform(-1, 1)
        elif viscosity <= 100:
            viscosity += random.uniform(-5, 5)
        else:
            viscosity += random.uniform(-10, 10)
        return round(viscosity, 1)

    def _get_cat(self, kind):
        for c in FQC_ITEMS.keys():
            if kind.find(c) == 0:
                return c

    def _get_col(self, item):
        col = ALL_FQC_ITEMS.index(item) + 5
        return col

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
        validity_date = ''
        ext_info = ''

        customer, internal_name, spec, batch, amount, product_date = [
            self._ws.cell(row=index, column=i).value for i in range(1, 7)]

        customer, internal_name, spec, batch, amount, product_date = map(null2str, [customer, internal_name,
                                                                                    spec, batch, amount, product_date])

        internal_name = re.sub(
            r'[\(\)（）]|20kg|20KG|5kg|5KG|1kg|1KG', ' ', internal_name)  # 去除不良字符
        internal_name = internal_name.strip()

        if is_number(batch):
            batch = str(int(batch))
        else:
            print("{}警告：Line {} 批号可能不是数字.{}".format(bcolors.WARNING,
                                                    index, bcolors.ENDC))
        if len(batch) != 8 and len(batch) != 6:
            ext_info += "(批号格式可能不正确)"
            print("{}Line{}:卧槽,批号格式不一般,小心地雷!!@#$@#&%{}".format(bcolors.FAIL,
                                                               index, bcolors.ENDC))

        if is_number(amount):
            # amount = '发货数量:{}kg'.format(int(amount))
            amount = ''
        else:
            amount = ''

        if not isinstance(product_date, datetime):
            print("{}Line{}:时间格式不正确,已经为你设置为空串.{}".format(bcolors.WARNING,
                                                         index, bcolors.ENDC))
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

            validity_date = datetime(
                year=next_year, month=next_month, day=next_day)

            product_date = datetime.strftime(product_date, '%Y/%m/%d')  # 转为字符串
            validity_date = datetime.strftime(
                validity_date, '%Y/%m/%d')  # 转为字符串

        return dict(customer=customer,
                    internal_name=internal_name,
                    spec=spec,
                    batch=batch,
                    amount=amount,
                    product_date=product_date,
                    validity_date=validity_date,
                    ext_info=ext_info)

    def _query_info(self, given):
        ''' 查询数据库,追加更多信息 '''
        products = db.search_product(given['internal_name'])
        if products.count() == 0:
            print('数据库中无记录,请输入命令. \n 添加条目：add \n 跳过此行：break \n 编辑字段：edit \n 退出：任意其他字符')
            cmd = rlinput("Command:")
            if cmd == "add":
                market_name = rlinput("销售品名:")
                template = rlinput("模板:")
                viscosity = rlinput("粘度:")
                viscosity_width = rlinput("粘度幅度:")
                product_obj = Product(internal_name=given['internal_name'],
                                      market_name=market_name,
                                      template=template,
                                      viscosity=int(viscosity),
                                      viscosity_width=int(viscosity_width))
                db.insert_product(product_obj)
                db.insert_product_to_xlsx(product_obj,
                                          '%s/data/database.xlsx' % self.app_path,
                                          'products')
                print("已经插入新的条目到products数据表")
            elif cmd == "break":
                return
            elif cmd == "edit":
                given['internal_name'] = rlinput("品名:", given['internal_name'])
                self._ws.cell('B{}'.format(self.index)).value = given[
                    'internal_name']
                try:
                    self._wb.save(self._product_wb_file)
                except PermissionError:
                    print("war:文件已经被打开，无法写入")
                return self._query_info(given)
            else:
                sys.exit()
        elif products.count() == 1:
            product_obj = products.one()
        else:
            ids = []
            for product in products.all():
                ids.append(product.id)
                space = " " * (20 - len(product.internal_name)
                               ) if len(product.internal_name) < 20 else ""
                print("\t %s%s\t ID:%s\t %s±%sdPa.s" % (product.internal_name,
                                                        space,
                                                        product.id,
                                                        product.viscosity,
                                                        product.viscosity_width))

            while True:
                print("小提示: 你可以输入quit立即退出")
                pid = rlinput("please choise a ID:")
                if pid == "quit":
                    sys.exit()
                if self._validate_id(pid, ids):
                    break

            product_obj = db.get_product_by_id(pid)
            print("\t 你选择了(%s, %s±%sdPa.s)" % (product_obj.internal_name,
                                               product_obj.viscosity,
                                               product_obj.viscosity_width))

        if product_obj.market_name.find('SP8') >= 0 or \
                product_obj.market_name == 'A-9060A 01':
            given['ext_info'] += '(深南电路要求打发货数量)'

        if product_obj.market_name.find('28GHB') >= 0 or \
                product_obj.market_name.find('30GHB') >= 0 or \
                product_obj.market_name.find('SP20HF') >= 0:
            given['ext_info'] += '(宏华胜要求打发货数量)'

        if product_obj.market_name == '8BL' or \
                product_obj.market_name == 'GH3' or \
                product_obj.market_name.find('G6') >= 0 or \
                product_obj.market_name.find('MG31') >= 0:
            given['ext_info'] += '(大连崇达要求打发货数量)'

        given['market_name'] = product_obj.market_name
        given['kind'] = product_obj.template
        given['template'] = TEMPLATES[product_obj.template]
        given['viscosity'] = product_obj.viscosity
        given['viscosity_limit'] = "%s±%s" % (product_obj.viscosity,
                                              product_obj.viscosity_width)
        given['qc_date'] = datetime.strftime(datetime.now(), '%Y/%m/%d')
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

    def get_template(self, name):
        ''' 获取模板文件路径 '''
        return os.path.join(self.app_path,
                            "templates/%s.docx" % name)


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
        db.init_product_data(
            '{}/data/database.xlsx'.format(module_path()), 'products')
    elif args.init_materials:
        reset_table('iqc_materials')
        db.init_materials(
            '{}/data/原材料检验项目及要求.xlsx'.format(module_path()), 'Sheet1')
    elif args.iqc:
        if args.filename:
            generate_iqc_reports(args.filename, args.end_row)
        else:
            print(
                "filename parameter must be provided. please refer --help for more detail.")
    else:
        print("unknown parameter, please refer --help for more detail.")
