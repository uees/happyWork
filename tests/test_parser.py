# -*- coding: utf-8 -*-
import os
import unittest

from settings import BASE_DIR
from formula.parser import FormulaParser


class FormulaParserTestCase(unittest.TestCase):

    def setUp(self):
        self.parser = FormulaParser(os.path.join(
            BASE_DIR,
            'data/工单/UV阻焊及文字产品/UVS系类/2012-09-18_UVS-1000 GY 黄绿(B-01).xlsx'
        ))

    def tearDown(self):
        self.parser = None

    def test_parse_filename(self):
        formula = self.parser.parse_filename()
        assert isinstance(formula, dict), '不匹配'
        self.assertEqual(formula['created_at'], '2012-09-18')
        self.assertEqual(formula['name'], 'UVS-1000 GY 黄绿')
        self.assertEqual(formula['version'], 'B-01')
        self.assertEqual(formula['description'], '')

    def test_get_products_in_sheet(self):
        products = self.parser.get_products()
        self.assertEqual(len(products), 1)
        for product_name, sheet in products:
            self.assertEqual(product_name, "UVS-1000 GY 黄绿")
            worksheet = self.parser.workbook[sheet]
            product_name2 = worksheet['C4'].value
            self.assertEqual(product_name2, None)

    def test_get_mixing_materials(self):
        worksheet = self.parser.workbook.worksheets[0]
        assert worksheet['A7'].value.startswith("树脂")  # 树脂区
        value = worksheet['C17'].value
        self.assertTrue(isinstance(value, float) or isinstance(value, int), '不是数字')

        materials = self.parser.get_mixing_materials()
        self.assertEqual(len(materials), 13, '有材料没采集到')
        self.assertEqual(materials[4]['name'], 'A0113', "不是A0113")

    def test_get_extends_info(self):
        worksheet = self.parser.workbook.worksheets[1]  # 生产单
        self.assertTrue(isinstance(worksheet['C16'].value, float))
        self.assertEqual(worksheet['C18'].value, None)
        self.assertTrue(isinstance(worksheet['C19'].value, str))
        self.assertTrue(isinstance(worksheet['A19'].value, str))

        info = self.parser.get_extends_info('生产单')
        self.assertEqual(len(info['materials']), 2)
        self.assertEqual(len(info['requirements']), 1)

    def test_get_formulas(self):
        formulas = self.parser.parse()
        for formula in formulas:
            print(formula)


if __name__ == "__main__":
    unittest.main()
