#!/usr/bin/env python
# -*- coding: utf-8 -*-

import unittest
import mock
import os
import sys
import yaml
import shutil
import datetime

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
        actually_result = ted.check_if_file_exists("../fielddefinition.yml")
        self.assertEqual(expected_result, actually_result)

    def test_case_02(self):
        expected_result = False
        actually_result = ted.check_if_file_exists("not_exist.yml")
        self.assertEqual(expected_result, actually_result)

###############################################################################
# test -> get_yaml_data
class Test_get_yaml_data(unittest.TestCase):

    def setUp(self):
        self.yaml_filename = "../fielddefinition.yml"

    def tearDown(self):
        pass

    def test_case_01(self):
        expected_result = dict
        actually_result = type(ted.get_yaml_data(self.yaml_filename))
        self.assertEqual(expected_result, actually_result)

    def test_case_02(self):
        expected_result = (
"""
read_target:
  file_name: "/mnt/c/Temp/テスト_read.xlsx"
  sheet_number: 1
  sheet_name: "Sheet1"
  column_definition: ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
  row_definition: 2
  row_max: 10

write_target:
  file_name: "/mnt/c/Temp/write.xlsx"
  sheet_number: 1
  sheet_name: "Sheet1"
  column_definition: ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
  row_definition: 4
"""     )
        actually_result = ted.get_yaml_data(self.yaml_filename)
        self.assertEqual(dict(yaml.safe_load(expected_result)), actually_result)

    def test_case_03(self):
        expected_result = FileNotFoundError
        with self.assertRaises(expected_result):
            yaml.load(ted.get_yaml_data("not_exist.yml"))

###############################################################################
# test -> readparent2writeparent
class Test_readparent2writeparent(unittest.TestCase):

    def setUp(self):
        pass

    def tearDown(self):
        pass

    def test_case_01(self):
        expected_result = "A7"
        actually_result = ted.readparent2writeparent("A5", 2, 4)
        self.assertEqual(expected_result, actually_result)

    def test_case_02(self):
        expected_result = "A5"
        actually_result = ted.readparent2writeparent("A5", 2, 2)
        self.assertEqual(expected_result, actually_result)

    def test_case_03(self):
        expected_result = "AA53"
        actually_result = ted.readparent2writeparent("AA51", 2, 4)
        self.assertEqual(expected_result, actually_result)

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
        td_now = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        template_file_name = "./template_write.xlsx"
        self.xlsx_file_name = "./result_{testname}_{time}.xlsx".format(
                                testname=self.__class__.__name__, time=td_now)
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
# test -> OpyxlWrapper -> get_cellalignment
class Test_get_cellalignment(unittest.TestCase):

    def setUp(self):
        self.ow = ted.OpyxlWrapper("テスト_read.xlsx")

    def tearDown(self):
        self.ow.close_workbook()

    # セルの書式設定 -> 配置 -> 縦位置 -> 中央揃え: center
    def test_case_01(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = "center"
        actually_result = self.ow.get_cellalignment(4, "B")
        self.assertEqual(expected_result, actually_result.vertical)
    # セルの書式設定 -> 配置 -> 横位置 -> 標準: 0
    ## 横位置を「標準」で定義する場合は "general" を指定するが、
    ## get_cellalignmentした結果は 0 となるため、ここでは「標準」を 0 として処理することとする。
    def test_case_02(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = None
        actually_result = self.ow.get_cellalignment(4, "B")
        self.assertEqual(expected_result, actually_result.horizontal)
    # セルの書式設定 -> 配置 -> 方向(文字の回転) -> 定義しない: 0
    def test_case_03(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = 0
        actually_result = self.ow.get_cellalignment(4, "B")
        self.assertEqual(expected_result, actually_result.text_rotation)
    # セルの書式設定 -> 配置 -> 折り返して全体を表示する -> チェックしない: None
    def test_case_04(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = None
        actually_result = self.ow.get_cellalignment(4, "B")
        self.assertEqual(expected_result, actually_result.wrap_text)
    # セルの書式設定 -> 配置 -> 縮小して全体を表示する -> チェックしない: None
    def test_case_05(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = None
        actually_result = self.ow.get_cellalignment(4, "B")
        self.assertEqual(expected_result, actually_result.shrink_to_fit)
    # セルの書式設定 -> 配置 -> インデント -> 0: 0
    def test_case_06(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = 0
        actually_result = self.ow.get_cellalignment(4, "B")
        self.assertEqual(expected_result, actually_result.indent)
    # セルの書式設定 -> 配置 -> 前後にスペースを入れる -> 定義不可(チェックしない): None
    def test_case_07(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = None
        actually_result = self.ow.get_cellalignment(4, "B")
        self.assertEqual(expected_result, actually_result.justifyLastLine)
    # セルの書式設定 -> 配置 -> 文字の方向 -> 0: 0
    def test_case_08(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = 0
        actually_result = self.ow.get_cellalignment(4, "B")
        self.assertEqual(expected_result, actually_result.readingOrder)

###############################################################################
# test -> OpyxlWrapper -> _is_mergedcell_correct_format
class Test_is_mergedcell_correct_format(unittest.TestCase):

    def setUp(self):
        pass

    def tearDown(self):
        pass
    # Trueのパターン: 番号001-099の番号にすること
    def test_case_001(self):
        expected_result = True
        actually_result = ted.OpyxlWrapper._is_mergedcell_correct_format(None, "C3:C4")
        self.assertEqual(expected_result, actually_result)

    def test_case_002(self):
        expected_result = True
        actually_result = ted.OpyxlWrapper._is_mergedcell_correct_format(None, "AC3:C4")
        self.assertEqual(expected_result, actually_result)

    def test_case_003(self):
        expected_result = True
        actually_result = ted.OpyxlWrapper._is_mergedcell_correct_format(None, "A3:CC4")
        self.assertEqual(expected_result, actually_result)

    def test_case_004(self):
        expected_result = True
        actually_result = ted.OpyxlWrapper._is_mergedcell_correct_format(None, "AA32:CC43")
        self.assertEqual(expected_result, actually_result)
    # Falseのパターン: 番号100以上の番号にすること
    def test_case_100(self):
        expected_result = False
        actually_result = ted.OpyxlWrapper._is_mergedcell_correct_format(None, "3A:C4")
        self.assertEqual(expected_result, actually_result)

    def test_case_101(self):
        expected_result = False
        actually_result = ted.OpyxlWrapper._is_mergedcell_correct_format(None, "A3:4C")
        self.assertEqual(expected_result, actually_result)

    def test_case_102(self):
        expected_result = False
        actually_result = ted.OpyxlWrapper._is_mergedcell_correct_format(None, "A3:C4A")
        self.assertEqual(expected_result, actually_result)

###############################################################################
# test -> OpyxlWrapper -> get_mergedcells_list
class Test_get_mergedcells_list(unittest.TestCase):

    def setUp(self):
        self.ow = None

    def tearDown(self):
        # close workbook
        self.ow.close_workbook()

    def test_case_01(self):
        self.xlsx_file_name = "./テスト_read.xlsx"
        self.ow = ted.OpyxlWrapper(self.xlsx_file_name)
        # set workbook
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = [
            ['C3','C4'],
            ['D3','D4'],
            ['F3','F4']
        ]
        actually_result = self.ow.get_mergedcells_list()
        self.assertEqual(expected_result, actually_result)

    def test_case_02(self):
        self.xlsx_file_name = "./template_write.xlsx"
        self.ow = ted.OpyxlWrapper(self.xlsx_file_name)
        # set workbook
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = []
        actually_result = self.ow.get_mergedcells_list()
        self.assertEqual(expected_result, actually_result)

###############################################################################
# test -> OpyxlWrapper -> merge_cells
class Test_merge_cells(unittest.TestCase):

    def setUp(self):
        td_now = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        template_file_name = "./template_write.xlsx"
        self.xlsx_file_name = "./result_{testname}_{time}.xlsx".format(
                                testname=self.__class__.__name__, time=td_now)
        # copy test file
        if os.path.exists(self.xlsx_file_name) == False:
            # copy template_write.xlsx to <xlsx_file_name>
            shutil.copyfile(template_file_name, self.xlsx_file_name)
        self.ow = ted.OpyxlWrapper(self.xlsx_file_name)

    def tearDown(self):
        # close workbook
        self.ow.close_workbook()

    def test_case_01(self):
        # set workbook
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        self.ow.merge_cells("B4:B6")
        self.ow.merge_cells("C4:C7")
        self.ow.merge_cells("F2:F8")
        self.ow.save_workbook()
        self.ow.close_workbook()
        # reopen workbook
        self.ow = ted.OpyxlWrapper(self.xlsx_file_name)
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        actually_result = self.ow.get_mergedcells_list()
        expected_result = [
            ["B4","B6"],
            ["C4","C7"],
            ["F2","F8"]
            ]
        self.assertEqual(expected_result, actually_result)


if __name__ == "__main__":
    unittest.main()
