#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import traceback
import yaml
import re

try:
    from transcribeexceldata import OpyxlWrapper
except:
    print("[ERROR] transcribeexceldata.py が配置されていない可能性があります。")
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

def readparent2writeparent(read_parent, read_row, write_row)->str:
    re_result = re.search(r'^([A-Z]+)([0-9]+)$', read_parent)
    if re_result == None:
        raise Exception("[ERROR readparent2writeparent] read_parent is incorrect format. -> " + str(read_parent))
    row_delta = write_row - read_row
    return re_result.group(1) + str(int(re_result.group(2)) + row_delta)

def main():
    print("[INFO] 処理を開始します。")
    try:
        # fielddefinition.yml を読み込む
        fielddefinition = "fielddefinition.yml"
        if check_if_file_exists(fielddefinition) == False:
            raise Exception("[ERROR main] ファイル {file} が存在しません。".format(file=fielddefinition))
        yaml_data = get_yaml_data(fielddefinition)

        # 読み込み対象のExcelファイルを読み込む
        read_target_filename = yaml_data["read_target"]["file_name"]
        if check_if_file_exists(read_target_filename) == False:
            raise Exception("[ERROR main] ファイル {file} が存在しません。".format(file=read_target_filename))
        read_target_row = yaml_data["read_target"]["row_definition"]
        read_target_row_max = yaml_data["read_target"]["row_max"]

        # 書き込み対象のExcelファイルを読み込む
        write_target_filename = yaml_data["write_target"]["file_name"]
        if check_if_file_exists(write_target_filename) == False:
            raise Exception("[ERROR main] ファイル {file} が存在しません。".format(file=write_target_filename))
        write_target_sheetnumber = yaml_data["write_target"].get("sheet_number", None)
        write_target_sheetname = yaml_data["write_target"].get("sheet_name", None)
        write_target_row = yaml_data["write_target"]["row_definition"]
        write_target_item_number_column = yaml_data["write_target"]["item_number_column"]
        opened_write_target = OpyxlWrapper(write_target_filename)
        opened_write_target.load_workbook()
        opened_write_target.load_worksheet(write_target_sheetnumber, write_target_sheetname)

        # 列を軸としてループする
        ## 指定されたカラムの分だけループ処理する
        # 書き込み対象の行の開始番号を定義する
        write_row = write_target_row
        write_target_row_max = read_target_row_max + (write_target_row - read_target_row)
        ## 項番を1から始める
        item_number = 1
        ## 行ループを開始する -> 処理する行の最大値までループする
        while write_row <= write_target_row_max:
            # 書き込み対象のカラムを指定する
            write_column = write_target_item_number_column
            # セルの値を取得する
            targetcell_value = opened_write_target.get_celldata(write_row, write_column)
            # セルの値が空でない場合のみセルへの書き込みを実行する
            if targetcell_value != None and targetcell_value != "":
                # セルの値(項番の番号)を書き込む
                opened_write_target.write_celldata(write_row, write_column, item_number)
                print("[INFO] 書き込みファイルのセル: {col}{row} を処理しました。項番: {item_num}".format(col=write_column, row=write_row, item_num=item_number))
                # 項番をインクリメントする
                item_number += 1
            # 行番号をインクリメントする
            write_row += 1

        # 保存して終了する
        opened_write_target.save_workbook()
        opened_write_target.close_workbook()
        print("[INFO] すべての処理が正常に終了しました。")
        sys.exit(0)

    except Exception as e:
        print("[ERROR] 処理が異常終了しました。")
        print(str(e))
        print(str(traceback.format_exc()))
        sys.exit(1)

if __name__ == "__main__":
    main()
