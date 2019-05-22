# -*- coding: utf-8 -*-
import os
import unittest

from settings import BASE_DIR
from formula.parser import FormulaParser


class FormulaAnalystTestCase(unittest.TestCase):

    def setUp(self):
        self.ast = FormulaParser(os.path.join(
            BASE_DIR, 'data/工单/UV阻焊及文字产品/UVS系类/2012-09-18_UVS-1000 GY 黄绿(B-01).xlsx'))

    def tearDown(self):
        self.ast = None

    def test_match_filemame(self):
        formula = self.ast.match_filemame()
        assert isinstance(formula, dict), '不匹配'
        self.assertEqual(formula['date'], '2015-10-22')
        self.assertEqual(formula['name'], 'UVS-1000 100dPa.s')
        self.assertEqual(formula['version'], 'B-02')
        self.assertEqual(formula['info'], '')

    def test_get_product_name_in_sheet(self):
        products = self.ast.get_product_name_in_sheet()
        self.assertTrue(len(products) == 1)
        name2 = self.ast.wb.get_sheet_by_name('生产单').cell('C4').value
        assert name2 is None, '不相等'

    def test_get_mixing_materials(self):
        assert self.ast.wb.worksheets[0].cell('A7').value == "树脂色浆区"
        value = self.ast.wb.worksheets[0].cell('C17').value
        self.assertTrue(isinstance(value, float) or isinstance(value, int), '不是数字')
        materials = self.ast.get_mixing_materials()
        self.assertTrue(len(materials) == 9, '有材料没采集到')
        self.assertEqual(materials[5][0], 'A0077', "应该是A0077才对")

    def test_get_jiaoliao_info(self):
        value = self.ast.wb.worksheets[1].cell('C17').value
        self.assertTrue(isinstance(value, float))
        value = self.ast.wb.worksheets[1].cell('C18').value
        self.assertTrue(isinstance(value, float))
        value = self.ast.wb.worksheets[1].cell('C19').value
        self.assertTrue(isinstance(value, str))
        value = self.ast.wb.worksheets[1].cell('C20').value
        self.assertTrue(value is None)
        value = self.ast.wb.worksheets[1].cell('A20').value
        self.assertTrue(isinstance(value, str))
        info = self.ast.get_jiaoliao_info('生产单')
        self.assertEqual(len(info['materials']), 4)
        self.assertEqual(len(info['yaoqiu']), 2)

    def test_get_formulas(self):
        formulas = self.ast.get_formulas()
        for formula in formulas:
            print(formula)


if __name__ == "__main__":
    unittest.main()
