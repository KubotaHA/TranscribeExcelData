#!/usr/bin/env python
# -*- coding: utf-8 -*-

import unittest
import mock
import os
import sys

sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
import mergecells as mc

###############################################################################
# test -> is_target_merge
class Test_is_target_merge(unittest.TestCase):

    def setUp(self):
        self.mergedcells_list = [
            ["A3","A5"],
            ["B4","B7"],
            ["C2","C8"],
            ["D1","D10"],
            ["D11","D13"]
        ]

    def tearDown(self):
        pass

    def test_case_01(self):
        write_target_reference_column = "A"
        actually_result_bool, actually_result_list = mc.is_target_merge(self.mergedcells_list, write_target_reference_column, 3,)
        self.assertEqual(True, actually_result_bool)
        self.assertEqual(["A3","A5"], actually_result_list)

    def test_case_02(self):
        write_target_reference_column = "A"
        expected_result = False
        actually_result_bool, actually_result_list = mc.is_target_merge(self.mergedcells_list, write_target_reference_column, 4,)
        self.assertEqual(False, actually_result_bool)
        self.assertEqual([], actually_result_list)

###############################################################################
# test -> merge_cells_parent
class Test_merge_cells_parent(unittest.TestCase):

    def setUp(self):
        pass

    def tearDown(self):
        pass

    def test_case_01(self):
        merged_cells = ["A3","A5"]
        write_column = "D"
        expected_result = "D3:D5"
        actually_result = mc.merge_cells_parent(merged_cells, write_column)
        self.assertEqual(expected_result, actually_result)

    def test_case_02(self):
        merged_cells = ["A3","A5"]
        write_column = "R"
        expected_result = "R3:R5"
        actually_result = mc.merge_cells_parent(merged_cells, write_column)
        self.assertEqual(expected_result, actually_result)

if __name__ == "__main__":
    unittest.main()
