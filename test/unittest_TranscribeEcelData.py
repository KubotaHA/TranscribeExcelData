#!/usr/bin/env python
# -*- coding: utf-8 -*-

import unittest
import mock
import os
import sys
import yaml

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
  file_name: "read.xlsx"
  sheet_number: 0
  sheet_name: "Sheet1"
  column_key: "B"
  column_definition: ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
  row_definition: 2

write_target:
  file_name: "write.xlsx"
  sheet_number: 0
  sheet_name: "Sheet1"
  column_definition: ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
  row_definition: 2
""")
        actually_result = ted.get_yaml_data("fielddefinition.yml")
        self.assertEqual(dict(yaml.load(expected_result)), actually_result)

    def test_case_03(self):
        expected_result = FileNotFoundError
        with self.assertRaises(expected_result):
            yaml.load(ted.get_yaml_data("not_exist.yml"))

if __name__ == "__main__":
    unittest.main()
