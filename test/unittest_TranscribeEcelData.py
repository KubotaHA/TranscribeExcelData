#!/usr/bin/env python
# -*- coding: utf-8 -*-

import unittest
import mock
import os
import sys
import yaml
import shutil
import time

sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
import transcribeexceldata as ted

###############################################################################
# test -> check_if_file_exists
class Test_check_if_file_exists(unittest.TestCase):

    def setUp(self):
        pass

    def tearDown(self):
        pass

    def test_case_01(self):
        expected_result = True
        actually_result = ted.check_if_file_exists("fielddefinition.yml")
        self.assertEqual(expected_result, actually_result)

    def test_case_02(self):
        expected_result = False
        actually_result = ted.check_if_file_exists("not_exist.yml")
        self.assertEqual(expected_result, actually_result)

###############################################################################
# test -> get_yaml_data
class Test_get_yaml_data(unittest.TestCase):

    def setUp(self):
        pass

    def tearDown(self):
        pass

    def test_case_01(self):
        expected_result = dict
        actually_result = type(ted.get_yaml_data("fielddefinition.yml"))
        self.assertEqual(expected_result, actually_result)

    def test_case_02(self):
        expected_result = (
"""
read_target:
  file_name: "テストread.xlsx"
  sheet_number: 1
  sheet_name: "Sheet1"
  column_key: "B"
  column_definition: ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
  row_definition: 2

write_target:
  file_name: "write.xlsx"
  sheet_number: 1
  sheet_name: "Sheet1"
  column_definition: ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
  row_definition: 4
""")
        actually_result = ted.get_yaml_data("fielddefinition.yml")
        self.assertEqual(dict(yaml.safe_load(expected_result)), actually_result)

    def test_case_03(self):
        expected_result = FileNotFoundError
        with self.assertRaises(expected_result):
            yaml.load(ted.get_yaml_data("not_exist.yml"))

###############################################################################
# test -> OpyxlWrapper -> get_cellcolor
class Test_get_cellcolor(unittest.TestCase):

    def setUp(self):
        self.ow = ted.OpyxlWrapper("テスト_read.xlsx")

    def tearDown(self):
        self.ow.close_workbook()

    # 黄色セル -> str: 'FFFFFF00'
    def test_case_01(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        actually_result = self.ow.get_cellcolor(2, "B")
        expected_result = 'FFFFFF00'
        self.assertEqual(expected_result, actually_result)
    # 黒セル -> int: 1
    def test_case_02(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        actually_result = self.ow.get_cellcolor(2, "G")
        expected_result = 1
        self.assertEqual(expected_result, actually_result)
    # 白セル -> int: 0
    def test_case_03(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        actually_result = self.ow.get_cellcolor(2, "H")
        expected_result = 0
        self.assertEqual(expected_result, actually_result)
    # 塗りつぶし無しセル -> str: '00000000'
    def test_case_04(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        actually_result = self.ow.get_cellcolor(2, "I")
        expected_result = '00000000'
        self.assertEqual(expected_result, actually_result)

###############################################################################
# test -> OpyxlWrapper -> fill_cellcolor
class Test_fill_cellcolor(unittest.TestCase):

    def setUp(self):
        template_file_name = "./template_write.xlsx"
        self.xlsx_file_name = "./write_fill_cellcolor_01.xlsx"
        # copy test file
        if os.path.exists(self.xlsx_file_name) == False:
            # copy template_write.xlsx to <xlsx_file_name>
            shutil.copyfile(template_file_name, self.xlsx_file_name)
        self.ow = ted.OpyxlWrapper(self.xlsx_file_name)

    def tearDown(self):
        # close workbook
        self.ow.close_workbook()

    # 黄色セル -> str: 'FFFFFF00'
    def test_case_01(self):
        # set workbook
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        self.ow.fill_cellcolor(4, "B", 'FFFFFF00')
        self.ow.save_workbook()
        self.ow.close_workbook()
        # reopen workbook
        self.ow = ted.OpyxlWrapper(self.xlsx_file_name)
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        actually_result = self.ow.get_cellcolor(4, "B")
        expected_result = 'FFFFFF00'
        self.assertEqual(expected_result, actually_result)
    # 白セル -> int: 1
    def test_case_02(self):
        # set workbook
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        self.ow.fill_cellcolor(4, "C", 1)
        self.ow.save_workbook()
        self.ow.close_workbook()
        # reopen workbook
        self.ow = ted.OpyxlWrapper(self.xlsx_file_name)
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        actually_result = self.ow.get_cellcolor(4, "C")
        expected_result = 1
        self.assertEqual(expected_result, actually_result)
    # 黒セル -> int: 0
    def test_case_03(self):
        # set workbook
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        self.ow.fill_cellcolor(4, "D", 0)
        self.ow.save_workbook()
        self.ow.close_workbook()
        # reopen workbook
        self.ow = ted.OpyxlWrapper(self.xlsx_file_name)
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        actually_result = self.ow.get_cellcolor(4, "D")
        expected_result = 0
        self.assertEqual(expected_result, actually_result)
    # 塗りつぶし無しセル -> str: '00000000'
    ## Colorクラスでは'00000000'は「黒」を示すが、色なしセルをget_cellcolorした返り値が'00000000'であるため、
    ## ここでは「色なし」を'00000000'として処理することとする。
    def test_case_04(self):
        # set workbook
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        self.ow.fill_cellcolor(4, "E", '00000000')
        self.ow.save_workbook()
        self.ow.close_workbook()
        # reopen workbook
        self.ow = ted.OpyxlWrapper(self.xlsx_file_name)
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        actually_result = self.ow.get_cellcolor(4, "E")
        expected_result = '00000000'
        self.assertEqual(expected_result, actually_result)

###############################################################################
# test -> get_alignment
class Test_get_alignment(unittest.TestCase):

    def setUp(self):
        self.ow = ted.OpyxlWrapper("テスト_read.xlsx")

    def tearDown(self):
        self.ow.close_workbook()

    # セルの書式設定 -> 配置 -> 縦位置 -> 中央揃え: center
    def test_case_01(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = "center"
        actually_result = self.ow.get_alignment(4, "B")
        self.assertEqual(expected_result, actually_result.vertical)
    # セルの書式設定 -> 配置 -> 横位置 -> 標準: 0
    ## 横位置を「標準」で定義する場合は "general" を指定するが、
    ## get_alignmentした結果は 0 となるため、ここでは「標準」を 0 として処理することとする。
    def test_case_02(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = None
        actually_result = self.ow.get_alignment(4, "B")
        self.assertEqual(expected_result, actually_result.horizontal)
    # セルの書式設定 -> 配置 -> 方向(文字の回転) -> 定義しない: 0
    def test_case_03(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = 0
        actually_result = self.ow.get_alignment(4, "B")
        self.assertEqual(expected_result, actually_result.text_rotation)
    # セルの書式設定 -> 配置 -> 折り返して全体を表示する -> チェックしない: None
    def test_case_04(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = None
        actually_result = self.ow.get_alignment(4, "B")
        self.assertEqual(expected_result, actually_result.wrap_text)
    # セルの書式設定 -> 配置 -> 縮小して全体を表示する -> チェックしない: None
    def test_case_05(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = None
        actually_result = self.ow.get_alignment(4, "B")
        self.assertEqual(expected_result, actually_result.shrink_to_fit)
    # セルの書式設定 -> 配置 -> インデント -> 0: 0
    def test_case_06(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = 0
        actually_result = self.ow.get_alignment(4, "B")
        self.assertEqual(expected_result, actually_result.indent)
    # セルの書式設定 -> 配置 -> 前後にスペースを入れる -> 定義不可(チェックしない): None
    def test_case_07(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = None
        actually_result = self.ow.get_alignment(4, "B")
        self.assertEqual(expected_result, actually_result.justifyLastLine)
    # セルの書式設定 -> 配置 -> 文字の方向 -> 0: 0
    def test_case_08(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = 0
        actually_result = self.ow.get_alignment(4, "B")
        self.assertEqual(expected_result, actually_result.readingOrder)

if __name__ == "__main__":
    unittest.main()
