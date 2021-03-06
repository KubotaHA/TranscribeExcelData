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
  file_name:                             "/mnt/c/Temp/ใในใ_read.xlsx"
  sheet_number:                          1
  #sheet_name:                            "Sheet1"
  column_definition:                     ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
  row_definition:                        2
  row_max:            &read_row_max      100

write_target:
  file_name:          &write_file        "/mnt/c/Temp/write.xlsx"
  sheet_number:       &write_sheet_num   1
  #sheet_name:         &write_sheet_name  "Sheet1"
  column_definition:                     ["B", "D", "C", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
  row_definition:     &write_row         4

numberitems_target:
  file_name:          *write_file
  sheet_number:       *write_sheet_num
  #sheet_name:         *write_sheet_name
  row_definition:     *write_row
  row_max:            *read_row_max
  item_number_column:                    "B"

merge_target:
  file_name:          *write_file
  sheet_number:       *write_sheet_num
  #sheet_name:         *write_sheet_name
  row_definition:     *write_row
  row_max:            *read_row_max
  target_column:                         ["E", "F", "G"]
  reference_column:                      "D"
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
        actually_result = ted.readparent2writeparent("A5", 2, 4, "A")
        self.assertEqual(expected_result, actually_result)

    def test_case_02(self):
        expected_result = "B5"
        actually_result = ted.readparent2writeparent("A5", 2, 2, "B")
        self.assertEqual(expected_result, actually_result)

    def test_case_03(self):
        expected_result = "C53"
        actually_result = ted.readparent2writeparent("AA51", 2, 4, "C")
        self.assertEqual(expected_result, actually_result)

###############################################################################
# test -> OpyxlWrapper -> get_cellcolor
class Test_get_cellcolor(unittest.TestCase):

    def setUp(self):
        self.ow = ted.OpyxlWrapper("ใในใ_read.xlsx")

    def tearDown(self):
        self.ow.close_workbook()

    # ้ป่ฒใปใซ -> str: 'FFFFFF00'
    def test_case_01(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        actually_result_colorindex, actually_result_index_filltype= self.ow.get_cellcolor(2, "B")
        self.assertEqual('FFFFFF00', actually_result_colorindex)
        self.assertEqual("solid", actually_result_index_filltype)
    # ้ปใปใซ -> int: 1
    def test_case_02(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        actually_result_colorindex, actually_result_index_filltype = self.ow.get_cellcolor(2, "G")
        self.assertEqual(1, actually_result_colorindex)
        self.assertEqual("solid", actually_result_index_filltype)
    # ็ฝใปใซ -> int: 0
    def test_case_03(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        actually_result_colorindex, actually_result_index_filltype = self.ow.get_cellcolor(2, "H")
        self.assertEqual(0, actually_result_colorindex)
        self.assertEqual("solid", actually_result_index_filltype)
    # ๅกใใคใถใ็กใใปใซ -> str: '00000000'
    def test_case_04(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        actually_result_colorindex, actually_result_index_filltype = self.ow.get_cellcolor(2, "I")
        self.assertEqual('00000000', actually_result_colorindex)
        self.assertEqual(None, actually_result_index_filltype)

###############################################################################
# test -> OpyxlWrapper -> fill_cellcolor
# class Test_fill_cellcolor(unittest.TestCase):

#     def setUp(self):
#         td_now = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
#         template_file_name = "./template_write.xlsx"
#         self.xlsx_file_name = "./result_{testname}_{time}.xlsx".format(
#                                 testname=self.__class__.__name__, time=td_now)
#         # copy test file
#         if os.path.exists(self.xlsx_file_name) == False:
#             # copy template_write.xlsx to <xlsx_file_name>
#             shutil.copyfile(template_file_name, self.xlsx_file_name)
#         self.ow = ted.OpyxlWrapper(self.xlsx_file_name)

#     def tearDown(self):
#         # close workbook
#         self.ow.close_workbook()

#     # ้ป่ฒใปใซ -> str: 'FFFFFF00'
#     def test_case_01(self):
#         # set workbook
#         self.ow.load_workbook()
#         self.ow.load_worksheet(1, None)
#         self.ow.fill_cellcolor(4, "B", 'FFFFFF00')
#         self.ow.save_workbook()
#         self.ow.close_workbook()
#         # reopen workbook
#         self.ow = ted.OpyxlWrapper(self.xlsx_file_name)
#         self.ow.load_workbook()
#         self.ow.load_worksheet(1, None)
#         actually_result_colorindex, actually_result_filltype = self.ow.get_cellcolor(4, "B")
#         self.assertEqual('FFFFFF00', actually_result_colorindex)
#         self.assertEqual('solid', actually_result_filltype)
#     # ็ฝใปใซ -> int: 1
#     def test_case_02(self):
#         # set workbook
#         self.ow.load_workbook()
#         self.ow.load_worksheet(1, None)
#         self.ow.fill_cellcolor(4, "C", 1)
#         self.ow.save_workbook()
#         self.ow.close_workbook()
#         # reopen workbook
#         self.ow = ted.OpyxlWrapper(self.xlsx_file_name)
#         self.ow.load_workbook()
#         self.ow.load_worksheet(1, None)
#         actually_result_colorindex, actually_result_filltype = self.ow.get_cellcolor(4, "C")
#         self.assertEqual(0, actually_result_colorindex)
#         self.assertEqual("solid", actually_result_filltype)
#     # ้ปใปใซ -> int: 0
#     def test_case_03(self):
#         # set workbook
#         self.ow.load_workbook()
#         self.ow.load_worksheet(1, None)
#         self.ow.fill_cellcolor(4, "D", 0)
#         self.ow.save_workbook()
#         self.ow.close_workbook()
#         # reopen workbook
#         self.ow = ted.OpyxlWrapper(self.xlsx_file_name)
#         self.ow.load_workbook()
#         self.ow.load_worksheet(1, None)
#         actually_result_colorindex, actually_result_filltype = self.ow.get_cellcolor(4, "D")
#         self.assertEqual(1, actually_result_colorindex)
#         self.assertEqual("solid", actually_result_filltype)
#     # ๅกใใคใถใ็กใใปใซ -> str: '00000000'
#     ## Colorใฏใฉในใงใฏ'00000000'ใฏใ้ปใใ็คบใใใ่ฒใชใใปใซใget_cellcolorใใ่ฟใๅคใ'00000000'ใงใใใใใ
#     ## ใใใงใฏใ่ฒใชใใใ filltype == None ใฎๅ?ดๅใฎใฟๅฆ็ใใใใจใจใใใ
#     def test_case_04(self):
#         # set workbook
#         self.ow.load_workbook()
#         self.ow.load_worksheet(1, None)
#         self.ow.fill_cellcolor(4, "E", '00000000', None)
#         self.ow.save_workbook()
#         self.ow.close_workbook()
#         # reopen workbook
#         self.ow = ted.OpyxlWrapper(self.xlsx_file_name)
#         self.ow.load_workbook()
#         self.ow.load_worksheet(1, None)
#         actually_result_colorindex, actually_result_filltype = self.ow.get_cellcolor(4, "E")
#         self.assertEqual("00000000", actually_result_colorindex)
#         self.assertEqual(None, actually_result_filltype)

###############################################################################
# test -> OpyxlWrapper -> get_cellalignment
class Test_get_cellalignment(unittest.TestCase):

    def setUp(self):
        self.ow = ted.OpyxlWrapper("ใในใ_read.xlsx")

    def tearDown(self):
        self.ow.close_workbook()

    # ใปใซใฎๆธๅผ่จญๅฎ -> ้็ฝฎ -> ็ธฆไฝ็ฝฎ -> ไธญๅคฎๆใ: center
    def test_case_01(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = "center"
        actually_result = self.ow.get_cellalignment(4, "B")
        self.assertEqual(expected_result, actually_result.vertical)
    # ใปใซใฎๆธๅผ่จญๅฎ -> ้็ฝฎ -> ๆจชไฝ็ฝฎ -> ๆจๆบ: 0
    ## ๆจชไฝ็ฝฎใใๆจๆบใใงๅฎ็พฉใใๅ?ดๅใฏ "general" ใๆๅฎใใใใ
    ## get_cellalignmentใใ็ตๆใฏ 0 ใจใชใใใใใใใงใฏใๆจๆบใใ 0 ใจใใฆๅฆ็ใใใใจใจใใใ
    def test_case_02(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = None
        actually_result = self.ow.get_cellalignment(4, "B")
        self.assertEqual(expected_result, actually_result.horizontal)
    # ใปใซใฎๆธๅผ่จญๅฎ -> ้็ฝฎ -> ๆนๅ(ๆๅญใฎๅ่ปข) -> ๅฎ็พฉใใชใ: 0
    def test_case_03(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = 0
        actually_result = self.ow.get_cellalignment(4, "B")
        self.assertEqual(expected_result, actually_result.text_rotation)
    # ใปใซใฎๆธๅผ่จญๅฎ -> ้็ฝฎ -> ๆใ่ฟใใฆๅจไฝใ่กจ็คบใใ -> ใใงใใฏใใชใ: None
    def test_case_04(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = None
        actually_result = self.ow.get_cellalignment(4, "B")
        self.assertEqual(expected_result, actually_result.wrap_text)
    # ใปใซใฎๆธๅผ่จญๅฎ -> ้็ฝฎ -> ็ธฎๅฐใใฆๅจไฝใ่กจ็คบใใ -> ใใงใใฏใใชใ: None
    def test_case_05(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = None
        actually_result = self.ow.get_cellalignment(4, "B")
        self.assertEqual(expected_result, actually_result.shrink_to_fit)
    # ใปใซใฎๆธๅผ่จญๅฎ -> ้็ฝฎ -> ใคใณใใณใ -> 0: 0
    def test_case_06(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = 0
        actually_result = self.ow.get_cellalignment(4, "B")
        self.assertEqual(expected_result, actually_result.indent)
    # ใปใซใฎๆธๅผ่จญๅฎ -> ้็ฝฎ -> ๅๅพใซในใใผในใๅฅใใ -> ๅฎ็พฉไธๅฏ(ใใงใใฏใใชใ): None
    def test_case_07(self):
        self.ow.load_workbook()
        self.ow.load_worksheet(1, None)
        expected_result = None
        actually_result = self.ow.get_cellalignment(4, "B")
        self.assertEqual(expected_result, actually_result.justifyLastLine)
    # ใปใซใฎๆธๅผ่จญๅฎ -> ้็ฝฎ -> ๆๅญใฎๆนๅ -> 0: 0
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
    # Trueใฎใใฟใผใณ: ็ชๅท001-099ใฎ็ชๅทใซใใใใจ
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
    # Falseใฎใใฟใผใณ: ็ชๅท100ไปฅไธใฎ็ชๅทใซใใใใจ
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
        self.xlsx_file_name = "./ใในใ_read.xlsx"
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

###############################################################################
# test -> check_if_duplicated
class Test_check_if_duplicated(unittest.TestCase):

    def setUp(self):
        pass

    def tearDown(self):
        pass

    def test_case_01(self):
        columnlist = ["A", "B", "C"]
        actually_result_bool, actually_result_col = ted.check_if_duplicated(columnlist)
        self.assertEqual(True, actually_result_bool)
        self.assertEqual(None, actually_result_col)

    def test_case_02(self):
        columnlist = ["A", "B", "B"]
        actually_result_bool, actually_result_col = ted.check_if_duplicated(columnlist)
        self.assertEqual(False, actually_result_bool)
        self.assertEqual("B", actually_result_col)

    def test_case_03(self):
        columnlist = ["A", "BB", "B", "BBB"]
        actually_result_bool, actually_result_col = ted.check_if_duplicated(columnlist)
        self.assertEqual(True, actually_result_bool)
        self.assertEqual(None, actually_result_col)

    def test_case_04(self):
        columnlist = ["A", "BB", "B", "BBB", "BB"]
        actually_result_bool, actually_result_col = ted.check_if_duplicated(columnlist)
        self.assertEqual(False, actually_result_bool)
        self.assertEqual("BB", actually_result_col)


if __name__ == "__main__":
    unittest.main()
