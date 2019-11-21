#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import pickle
import random
import re
import shutil

import win32com.client
from pywintypes import com_error

FLAGS = {
    'FQC阻焊表格': 5,
    'FQC湿膜表格': 5,
    'FQC其他油墨表格': 5,
}


class Fixer(object):
    """
    优化检测数据，供客户欣赏
    """

    def __init__(self, configfile='config'):

        self.configfile = configfile
        self.engine = win32com.client.Dispatch('Excel.Application')
        self.engine.Visible = False
        self.engine.DisplayAlerts = False

        self.filefrom = os.path.join(os.path.abspath(os.path.dirname(__file__)),
                                     'FQC检测记录表格.xlsx')
        self.fileto = 'E:\\品质部\\检测记录.xlsx'

        if not os.path.exists(self.fileto):
            shutil.copy(self.filefrom, self.fileto)
            self.flags = FLAGS
            self.wb = self.engine.Workbooks.Open(self.fileto)
        else:
            try:
                self.wb = self.engine.Workbooks(self.fileto)
            except com_error:
                self.wb = self.engine.Workbooks.Open(self.fileto)

            self.flags = self.load_flags()
            self.copy()

    def load_flags(self):
        if os.path.exists(self.configfile):
            with open(self.configfile, 'rb') as fp:
                flags = pickle.load(fp)
        else:
            flags = FLAGS
        return flags

    def save_flags(self, config):
        # pickle.dump(config, self.configfile)
        with open(self.configfile, 'wb') as fp:
            fp.write(pickle.dumps(config))

    # def copy(self):
    #    try:
    #        shutil.copy(self.filefrom, self.fileto)
    #    except PermissionError:
    #        self.engine.Workbooks(self.fileto).Close(-1)
    #        shutil.copy(self.filefrom, self.fileto)

    def copy(self):
        """ 执行数据拷贝，原样拷贝 """
        try:
            wb_from = self.engine.Workbooks(self.filefrom)
        except com_error:
            wb_from = self.engine.Workbooks.Open(self.filefrom)

        for sheetname, start in self.flags.items():
            ws_from = wb_from.Worksheets(sheetname)
            ws_to = self.wb.Worksheets(sheetname)

            max_row = ws_from.UsedRange.Rows.Count

            # 获取30列的数据
            data = ws_from.Range(ws_from.Cells(start, 1),
                                 ws_from.Cells(max_row, 30)).Value

            self.set_range_value(ws_to, 1, start, data)

        # 关闭 wb_from，防止意外打开
        wb_from.Close(-1)

    def set_range_value(self, ws, leftCol, topRow, data):
        """
        Insert a 2d array starting at given location.
        i.e. [['a','b','c'],['a','b','c'],['a','b','c']]
        Works out the size needed for itself
        """
        bottomRow = topRow + len(data) - 1
        rightCol = leftCol + len(data[0]) - 1
        ws.Range(ws.Cells(topRow, leftCol),
                 ws.Cells(bottomRow, rightCol)).Value = data

    def run(self, config):
        print('生成中，请不要关闭此窗口')
        starts = {}
        for sheetname, settings in config.items():
            ws = self.wb.Worksheets(sheetname)
            min_row = self.flags.get(sheetname)
            max_row = ws.UsedRange.Rows.Count
            print('%s: min_row=%s, max_row=%s' % (sheetname, min_row, max_row))

            new_start = min_row
            for row in range(min_row, max_row):
                for item in settings:
                    cell = ws.Cells(row, item['col'])
                    if 'done_col' in item:
                        done_cell = ws.Cells(row, item['done_col'])
                        if done_cell.Value:
                            new_start = row
                        getattr(self, item['func'])(cell, done_cell, *item['args'])
                    else:
                        getattr(self, item['func'])(cell, *item['args'])
            starts[sheetname] = new_start
        self.save_flags(starts)

    def fix(self, cell, goodvalue):
        if cell.Value and cell.Value != goodvalue:
            cell.Value = goodvalue
            cell.Font.ColorIndex = 1  # 黑色
            cell.Interior.ColorIndex = 2  # 白色

    def stuff(self, cell, done_cell, goodvalue):
        if done_cell.Value:
            if isinstance(goodvalue, list):
                cell.Value = random.choice(goodvalue)
            else:
                cell.Value = goodvalue

    def fix_xidu(self, cell, goodvalue):
        if cell.Value:
            try:
                value = float(cell.Value)
            except Exception:
                value = 21
            if value > 20:
                if isinstance(goodvalue, list):
                    cell.Value = random.choice(goodvalue)
                else:
                    cell.Value = goodvalue
                cell.Font.ColorIndex = 1  # 黑色
                cell.Interior.ColorIndex = 2  # 白色

    def fix_hongwai(self, cell, goodvalue=99.87):
        if cell.Value:
            try:
                value = float(cell.Value)
            except Exception:
                value = 98
            if value < 99:
                cell.Value = goodvalue if goodvalue else round(
                    random.uniform(99.3, 100), 2)
                cell.Font.ColorIndex = 1  # 黑色
                cell.Interior.ColorIndex = 2  # 白色

    def fix_ganguang(self, cell, guanglist):
        pattern = re.compile(r'^(\d+).*$')
        if cell.Value:
            match = pattern.match(cell.Value)
            if match:
                guang = int(match.groups()[0])
            else:
                guang = random.choice(guanglist)
            if guang not in guanglist:
                guang = random.choice(guanglist)
            cell.Value = "{}级".format(guang)

    def save(self, filename=None):
        """
        if filename is None, save the openning file
        else save as another file used the given name
        """
        if filename:
            self.wb.SaveAs(filename)
        else:
            self.wb.Save()

    def close(self):
        if self.wb:
            self.wb.Close(-1)
        if len(self.engine.Workbooks) == 0:
            self.engine.Quit()

    def protect(self, sheetname=None):
        if sheetname:
            self.wb.Worksheets(sheetname).Protect(Password='pzb')
        else:
            self.wb.Protect(Password='pzb')


if __name__ == "__main__":
    f = Fixer('config')
    config = {
        'FQC阻焊表格': [
            {
                'col': 8,   # 颜色
                'func': 'fix',
                'args': ['PASS']
            },
            {
                'col': 9,   # 细度
                'func': 'fix_xidu',
                'args': [[15, 17.5, 20]]
            },
            {
                'col': 12,   # 红外
                'func': 'fix_hongwai',
                'args': []
            },
            {
                'col': 13,   # 板面效果
                'func': 'fix',
                'args': ['PASS']
            },
            {
                'col': 14,   # 固化性
                'func': 'fix',
                'args': ['PASS']
            },
            {
                'col': 15,    # 感光性
                'func': 'fix_ganguang',
                'args': [range(9, 13)]
            },
            {
                'col': 23,  # 硬度
                'done_col': 18,   # 80 min 显影标志着做完
                'func': 'stuff',
                'args': ['6H']
            },
            {
                'col': 24,  # 附着力
                'done_col': 18,
                'func': 'stuff',
                'args': ['100%']
            },
            {
                'col': 25,   # 耐焊性
                'done_col': 18,
                'func': 'stuff',
                'args': ['PASS']
            },
            {
                'col': 26,   # 耐化性
                'done_col': 18,
                'func': 'stuff',
                'args': ['PASS']
            },
        ],
        'FQC湿膜表格': [
            {
                'col': 6,   # 颜色
                'func': 'fix',
                'args': ['PASS']
            },
            {
                'col': 7,   # 细度
                'func': 'fix_xidu',
                'args': [[15, 17.5, 20]]
            },
            {
                'col': 8,   # 红外
                'func': 'fix_hongwai',
                'args': []
            },
            {
                'col': 9,   # 板面效果
                'func': 'fix',
                'args': ['PASS']
            },
            {
                'col': 10,   # 固化性
                'func': 'fix',
                'args': ['PASS']
            },
            {
                'col': 11,   # 硬度
                'done_col': 16,  # 显影完成 标志着做完
                'func': 'stuff',
                'args': [['1H', '2H']]
            },
            {
                'col': 15,    # 感光性
                'func': 'fix_ganguang',
                'args': [range(6, 9)]
            },
        ],
        'FQC其他油墨表格': [
            {
                'col': 6,   # 颜色
                'func': 'fix',
                'args': ['PASS']
            },
            {
                'col': 7,   # 细度
                'func': 'fix_xidu',
                'args': [[15, 17.5, 20]]
            },
            {
                'col': 8,   # 红外
                'func': 'fix_hongwai',
                'args': []
            },
            {
                'col': 9,   # 板面效果
                'func': 'fix',
                'args': ['PASS']
            },
            {
                'col': 10,   # 固化性
                'done_col': 17,
                'func': 'stuff',
                'args': ['PASS']
            },
            {
                'col': 14,   # 去膜性
                'func': 'fix',
                'args': ['PASS']
            },
        ]
    }
    f.run(config)
    f.save()
    f.close()
