#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import traceback
import yaml

try:
    import openpyxl as opyxl
    from openpyxl.styles import PatternFill, Alignment
    from openpyxl.styles.colors import Color
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
        with open(yaml_file, encoding='utf8') as file:
            yaml_data = yaml.safe_load(file)
        return yaml_data
    except PermissionError as pe:
        raise PermissionError("[ERROR get_yaml_data] ファイルへのアクセス権限がありません。-> " + str(pe))
    except FileNotFoundError as fe:
        raise FileNotFoundError("[ERROR] get_yaml_data ファイルが存在しません。-> " + str(fe))
    except Exception as e:
            raise Exception("[ERROR get_yaml_data] Unexpected Error has been ocurred. -> " + str(e))

class OpyxlWrapper:
    def __init__(self, exel_file) -> None:
        self.exel_file = exel_file
        self.workbook = None
        self.worksheet = None

    def load_workbook(self):
        try:
            self.workbook = opyxl.load_workbook(filename=self.exel_file)
        except Exception as e:
            raise Exception("[ERROR load_workbook] Unexpected Error has been ocurred. -> " + str(e))
        return self.workbook

    # read_only でworkbookを読み込むとセルの一部の属性が取得できなくなるため基本的には使用しないこととする
    # def load_workbook_readonly(self):
    #     try:
    #         self.workbook = opyxl.load_workbook(filename=self.exel_file, read_only=True)
    #     except Exception as e:
    #         raise Exception("[ERROR load_workbook_readonly] Unexpected Error has been ocurred. -> " + str(e))

    def load_worksheet(self, sheet_number=None, sheet_name=None)->object:
        if self.workbook == None:
            raise Exception("[ERROR: load_worksheet] Exel file has not been loaded yet.")
        try:
            # シート番号で指定
            if sheet_number != None and sheet_name == None:
                # シート番号の指定が0以下はエラーとする
                if sheet_number < 1:
                    raise Exception("[ERROR load_worksheet] Unexpected sheet number. [number > 0]")
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
        except Exception as e:
            raise Exception("[ERROR load_worksheet] Unexpected Error has been ocurred. -> " + str(e))

    def get_celldata(self, row_number, column_a):
        if self.worksheet == None:
            raise Exception("[ERROR: get_celldata] Exel file and worksheet have not been loaded yet.")
        try:
            specified_cell = column_a + str(row_number)
            return self.worksheet[specified_cell].value
        except Exception as e:
            raise Exception("[ERROR get_celldata] Unexpected Error has been ocurred. -> " + str(e))

    def get_cellcolor(self, row_number, column_a)->str or int:
        if self.worksheet == None:
            raise Exception("[ERROR: get_cellcolor] Exel file and worksheet have not been loaded yet.")
        try:
            specified_cell = column_a + str(row_number)
            return self.worksheet[specified_cell].fill.start_color.index
        except Exception as e:
            raise Exception("[ERROR get_cellcolor] Unexpected Error has been ocurred. -> " + str(e))

    def get_alignment(self, row_number, column_a)->dict:
        if self.worksheet == None:
            raise Exception("[ERROR: get_alignment] Exel file and worksheet have not been loaded yet.")
        try:
            specified_cell = column_a + str(row_number)
            return self.worksheet[specified_cell].alignment
        except Exception as e:
            raise Exception("[ERROR get_alignment] Unexpected Error has been ocurred. -> " + str(e))

    def write_celldata(self, row_number, column_a, cell_value):
        if self.worksheet == None:
            raise Exception("[ERROR: write_celldata] Exel file and worksheet have not been loaded yet.")
        try:
            specified_cell = column_a + str(row_number)
            self.worksheet[specified_cell].value = cell_value
        except Exception as e:
            raise Exception("[ERROR write_celldata] Unexpected Error has been ocurred. -> " + str(e))

    # https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.fills.html
    # https://openpyxl.readthedocs.io/en/stable/styles.html
    def fill_cellcolor(self, row_number, column_a, color_index):
        try:
            if self.worksheet == None:
                raise Exception("Exel file and worksheet have not been loaded yet.")
            # color_index: 0-63
            if type(color_index) == int:
                my_color = Color(indexed=color_index)
            # color_index: 00000000-FFFFFFFF
            elif type(color_index) == str:
                my_color = Color(rgb=color_index)
            else:
                raise Exception("[ERROR: fill_cellcolor] Unexpected color_index has been ocurred..")
            specified_cell = column_a + str(row_number)
            # Colorクラスでは'00000000'は「黒」を示すが、色なしセルをget_cellcolorした返り値が'00000000'であるため、
            # ここでは「色なし」を'00000000'として処理することとする。
            # また、「黒」のセルについてget_cellcolorした場合の返り値は int型で 0 であるため、
            # fill_cellcolorの color_index に int型で 0 を指定することでfill可能である
            if color_index == '00000000':
                self.worksheet[specified_cell].fill = PatternFill(fill_type=None)
            else:
                self.worksheet[specified_cell].fill = PatternFill(patternType='solid', fgColor=my_color)
        except Exception as e:
            raise Exception("[ERROR get_cellcollor] Unexpected Error has been ocurred. -> " + str(e))

    # https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.alignment.html?highlight=Alignment#openpyxl.styles.alignment.Alignment
    # https://gammasoft.jp/blog/text-center-alignment-with-openpyxl/#openpyxl-alignment
    # https://cercopes-z.com/Python/openpyxl/module-styles-alignment-opxl.html
    def set_alignment(self, row_number, column_a, horizontal, vertical, text_rotation, wrap_text, shrink_to_fit, indent, justifyLastLine, readingOrder):
        if self.worksheet == None:
            raise Exception("[ERROR: set_alignment] Exel file and worksheet have not been loaded yet.")
        try:
            specified_cell = column_a + str(row_number)
            # 横位置を「標準」で定義する場合は "general" を指定するが、
            # get_alignmentした結果は 0 となるため、ここでは「標準」を 0 として処理することとする。
            if horizontal==None:
                horizontal="general"
            self.worksheet[specified_cell].alignment = Alignment(
                                                            horizontal=horizontal,
                                                            vertical=vertical,
                                                            text_rotation=text_rotation,
                                                            wrap_text=wrap_text,
                                                            shrink_to_fit=shrink_to_fit,
                                                            indent=indent,
                                                            justifyLastLine=justifyLastLine,
                                                            readingOrder=readingOrder
                                                            )
        except Exception as e:
            raise Exception("[ERROR set_alignment] Unexpected Error has been ocurred. -> " + str(e))

    def save_workbook(self):
        if self.workbook == None:
            raise Exception("[ERROR: save_workbook] Exel file has not been loaded yet.")
        try:
            self.workbook.save(self.exel_file)
        except Exception as e:
            raise Exception("[ERROR save_workbook] Unexpected Error has been ocurred. -> " + str(e))

    # https://openpyxl.readthedocs.io/en/stable/api/openpyxl.workbook.workbook.html?highlight=close#openpyxl.workbook.workbook.Workbook.close
    def close_workbook(self):
        if self.workbook == None:
            raise Exception("[ERROR: close_workbook] Exel file has not been loaded yet.")
        try:
            self.workbook.close()
            del self.exel_file
            del self.worksheet
            del self.workbook
        except Exception as e:
            raise Exception("[ERROR close_workbook] Unexpected Error has been ocurred. -> " + str(e))

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
        opened_read_target.load_workbook()
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
        # 行を軸としてループする -> 無限ループ防止処理は終わりの行で行う
        while True:
            # キーとなるカラムの値が空であることを検知したら処理を停止する
            key_value = opened_read_target.get_celldata(read_row, read_target_key)
            if key_value == None or key_value == "":
                break
            # 指定されたカラムの分だけ処理する
            for column_num in range(len(read_target_column)):
                # 読み込み対象のセルを指定する
                read_column = read_target_column[column_num]
                ## セルの値を取得する
                read_value = opened_read_target.get_celldata(read_row, read_column)
                ## セルの色を取得する
                read_color = opened_read_target.get_cellcolor(read_row, read_column)
                ## セルの書式設定を取得する
                read_alignment = opened_read_target.get_alignment(read_row, read_column)
                
                # 書き込み対象のカラムを指定する
                write_column = write_target_column[column_num]
                ## セルの値を書き込む
                opened_write_target.write_celldata(write_row, write_column, read_value)
                ## セルに色を適用する
                opened_write_target.fill_cellcolor(write_row, write_column, read_color)
                ## セルの書式設定を適用する
                opened_write_target.set_alignment(
                    write_row, write_column,
                    horizontal=read_alignment.horizontal,
                    vertical=read_alignment.vertical,
                    text_rotation=read_alignment.text_rotation,
                    wrap_text=read_alignment.wrap_text,
                    shrink_to_fit=read_alignment.shrink_to_fit,
                    indent=read_alignment.indent,
                    justifyLastLine=read_alignment.justifyLastLine,
                    readingOrder=read_alignment.readingOrder
                    )
                print("[INFO] 書き込みファイルのセル: {col}{row} を処理しました。".format(col=write_column, row=write_row))
            # 行番号をインクリメントする
            read_row += 1
            write_row += 1
            # 無限ループ防止 -> 10000行
            if read_row == 10000:
                break

        # 保存して終了する
        opened_read_target.close_workbook()
        opened_write_target.save_workbook()
        opened_write_target.close_workbook()
        print("[INFO] すべての処理が正常に終了しました。")
        sys.exit(0)

    except Exception as e:
        print(str(e))
        print(str(traceback.format_exc()))
        sys.exit(1)

if __name__ == "__main__":
    main()
