#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import traceback
import yaml

try:
    import openpyxl as opyxl
except:
    print("[ERROR] openpyxl がインストールされていない可能性があります。")
    print(traceback.format_exc())

def check_if_file_exists(file_path)->bool:
    if os.path.exists(file_path):
        return True
    else:
        return False

def get_yaml_data(yaml_file)->dict:
    try:
        with open(yaml_file) as file:
            yaml_data = yaml.safe_load(file)
        return yaml_data
    except PermissionError as pe:
        raise PermissionError("[ERROR get_yaml_data] ファイルへのアクセス権限がありません。-> " + str(pe))
    except FileNotFoundError as fe:
        raise FileNotFoundError("[ERROR] get_yaml_data ファイルが存在しません。-> " + str(fe))
    except Exception as e:
            raise Exception("[ERROR get_yaml_data] Unexpected Error has beec ocurred. -> " + str(e))

class OpyxlWrapper:
    def __init__(self, exel_file) -> None:
        self.exel_file = exel_file
        self.workbook = None
        self.worksheet = None

    def load_workbook(self):
        try:
            self.workbook = opyxl.load_workbook(filename=self.exel_file)
        except Exception as e:
            raise Exception("[ERROR load_workbook] Unexpected Error has beec ocurred. -> " + str(e))
        return self.workbook

    def load_workbook_readonly(self):
        try:
            self.workbook = opyxl.load_workbook(filename=self.exel_file, read_only=True)
        except Exception as e:
            raise Exception("[ERROR load_workbook_readonly] Unexpected Error has beec ocurred. -> " + str(e))
        return self.workbook


    def load_worksheet(self, sheet_number=None, sheet_name=None):
        if self.workbook == None:
            raise Exception("[ERROR: load_worksheet] Exel file has not been loaded yet.")
        try:
            # シート番号で指定
            if sheet_number != None and sheet_name == None:
                # シート番号の指定が0以下はエラーとする
                if sheet_number < 1:
                    raise Exception("[ERROR load_worksheet] Unexpected sheet number. -> " + str(e))
                worksheet_title = self.workbook.worksheets[sheet_number-1].title
                self.worksheet = self.workbook[worksheet_title]
            # シート名で指定
            elif sheet_number == None and sheet_name != None:
                self.worksheet = self.workbook[sheet_name]
            # シート名で指定 ※シート番号よりも優先
            elif sheet_number != None and sheet_name != None:
                self.worksheet = self.workbook[sheet_name]
            else:
                raise Exception("[ERROR: load_worksheet] sheet_number and sheet_name are None.")
            return self.worksheet
        except Exception as e:
            raise Exception("[ERROR load_worksheet] Unexpected Error has beec ocurred. -> " + str(e))

    def get_celldata(self, row_number, column_a):
        if self.worksheet == None:
            raise Exception("[ERROR: get_celldata] Exel file and worksheet have not been loaded yet.")
        try:
            specified_cell = column_a + str(row_number)
            return self.worksheet[specified_cell].value
        except Exception as e:
            raise Exception("[ERROR get_celldata] Unexpected Error has beec ocurred. -> " + str(e))

    def write_celldata(self, row_number, column_a, cell_value):
        if self.worksheet == None:
            raise Exception("[ERROR: write_celldata] Exel file and worksheet have not been loaded yet.")
        try:
            specified_cell = column_a + str(row_number)
            self.worksheet[specified_cell].value = cell_value
            return self.worksheet[specified_cell].value
        except Exception as e:
            raise Exception("[ERROR write_celldata] Unexpected Error has beec ocurred. -> " + str(e))

    def save_workbook(self):
        if self.workbook == None:
            raise Exception("[ERROR: save_workbook] Exel file has not been loaded yet.")
        try:
            self.workbook.save(self.exel_file)
        except Exception as e:
            raise Exception("[ERROR save_workbook] Unexpected Error has beec ocurred. -> " + str(e))

    def close_workbook(self):
        if self.workbook == None:
            raise Exception("[ERROR: close_workbook] Exel file has not been loaded yet.")
        try:
            self.workbook.close()
            self.workbook = None
        except Exception as e:
            raise Exception("[ERROR close_workbook] Unexpected Error has beec ocurred. -> " + str(e))

def main():
    try:
        # fielddefinition.yml を読み込む
        fielddefinition = "fielddefinition.yml"
        if check_if_file_exists(fielddefinition) == False:
            raise Exception("[ERROR main] ファイル {file} が存在しません。".format(file=fielddefinition))
        yaml_data = get_yaml_data(fielddefinition)

        # 転記元Excelファイルを読み込む
        read_target_filename = yaml_data["read_target"]["file_name"]
        if check_if_file_exists(read_target_filename) == False:
            raise Exception("[ERROR main] ファイル {file} が存在しません。".format(file=read_target_filename))
        read_target_sheetnumber = yaml_data["read_target"].get("sheet_number", None)
        read_target_sheetname = yaml_data["read_target"].get("sheet_name", None)
        read_target_column = yaml_data["read_target"]["column_definition"]
        read_target_row = yaml_data["read_target"]["row_definition"]
        read_target_key = yaml_data["read_target"]["column_key"]
        opened_read_target = OpyxlWrapper(read_target_filename)
        opened_read_target.load_workbook_readonly()
        opened_read_target.load_worksheet(read_target_sheetnumber, read_target_sheetname)

        # 転記先Excelファイルを読み込む
        write_target_filename = yaml_data["write_target"]["file_name"]
        if check_if_file_exists(write_target_filename) == False:
            raise Exception("[ERROR main] ファイル {file} が存在しません。".format(file=write_target_filename))
        write_target_sheetnumber = yaml_data["write_target"].get("sheet_number", None)
        write_target_sheetname = yaml_data["write_target"].get("sheet_name", None)
        write_target_column = yaml_data["write_target"]["column_definition"]
        write_target_row = yaml_data["write_target"]["row_definition"]
        opened_write_target = OpyxlWrapper(write_target_filename)
        opened_write_target.load_workbook()
        opened_write_target.load_worksheet(write_target_sheetnumber, write_target_sheetname)

        # 転記を実行する
        if len(read_target_column) != len(write_target_column):
            print("[ERROR main] 転記先と転記元における対象カラムの数が合いません。")
        read_row = read_target_row
        write_row = write_target_row
        while True:
            # キーとなるカラムの値が空であることを検知したら処理を停止する
            key_value = opened_read_target.get_celldata(read_row, read_target_key)
            if key_value == None or key_value == "":
                break
            for column_num in range(len(read_target_column)):
                read_column = read_target_column[column_num]
                read_value = opened_read_target.get_celldata(read_row, read_column)
                write_column = write_target_column[column_num]
                opened_write_target.write_celldata(write_row, write_column, read_value)
            read_row += 1
            write_row += 1
            # 無限ループ防止 -> 10000行
            if read_row == 10000:
                break

        # 保存して終了する
        opened_read_target.close_workbook()
        opened_write_target.save_workbook()
        opened_write_target.close_workbook()
        print("[INFO] 処理が正常に終了しました。")
        sys.exit(0)

    except Exception as e:
        print(str(e))
        print(str(traceback.format_exc()))
        sys.exit(1)

if __name__ == "__main__":
    main()
