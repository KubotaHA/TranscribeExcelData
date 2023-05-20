#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import traceback
import yaml
import re

try:
    from transcribeexceldata import OpyxlWrapper
    from transcribeexceldata import check_if_file_exists
    from transcribeexceldata import get_yaml_data
except:
    print("[ERROR] transcribeexceldata.py が配置されていない可能性があります。")
    print(traceback.format_exc())

def is_target_merge(mergedcells_list, write_target_reference_column, write_row):
    for merged_cells in mergedcells_list:
        if merged_cells[0] == (write_target_reference_column + str(write_row)):
            return True, merged_cells
    return False, []

def merge_cells_parent(merged_cells, write_column):
    re_result_start = re.search(r'^([A-Z]+)([0-9]+)$', merged_cells[0])
    re_result_end = re.search(r'^([A-Z]+)([0-9]+)$', merged_cells[1])
    if re_result_start == None:
        raise Exception("[ERROR is_target_merge] re_result_start is incorrect format. -> '{0}'".format(merged_cells[0]))
    if re_result_end == None:
        raise Exception("[ERROR is_target_merge] re_result_end is incorrect format. -> '{0}'".format(merged_cells[0]))
    mergerow_start = int(re_result_start.group(2))
    mergerow_end = int(re_result_end.group(2))
    return "{start_col}{start_row}:{end_col}{end_row}".format(
        start_col=write_column,
        start_row=mergerow_start,
        end_col=write_column,
        end_row=mergerow_end
    )

def main():
    print("[INFO] 処理を開始します。")
    try:
        # fielddefinition.yml を読み込む
        fielddefinition = "fielddefinition.yml"
        if check_if_file_exists(fielddefinition) == False:
            raise Exception("[ERROR main] ファイル {file} が存在しません。".format(file=fielddefinition))
        yaml_data = get_yaml_data(fielddefinition)

        # 書き込み対象のExcelファイルを読み込む
        write_target_filename = yaml_data["merge_target"]["file_name"]
        if check_if_file_exists(write_target_filename) == False:
            raise Exception("[ERROR main] ファイル {file} が存在しません。".format(file=write_target_filename))
        write_target_sheetnumber = yaml_data["merge_target"].get("sheet_number", None)
        write_target_sheetname = yaml_data["merge_target"].get("sheet_name", None)
        write_target_row = yaml_data["merge_target"]["row_definition"]
        write_target_row_max = yaml_data["merge_target"]["row_max"]
        write_target_target_column = yaml_data["merge_target"]["target_column"]
        write_target_reference_column = yaml_data["merge_target"]["reference_column"]
        # Excelファイルをロードする
        opened_write_target = OpyxlWrapper(write_target_filename)
        opened_write_target.load_workbook()
        opened_write_target.load_worksheet(write_target_sheetnumber, write_target_sheetname)
        # 対象のExcelファイルから結合されたセルのリストを取得する
        mergedcells_list = opened_write_target.get_mergedcells_list()

        # 列を軸としてループする
        ## 指定されたカラムの分だけループ処理する
        # 書き込み対象のカラムでループする
        for write_column in write_target_target_column:
            # 書き込み対象の行の開始番号を定義する
            write_row = write_target_row
            # 行ループを開始する -> 処理する行の最大値までループする
            while write_row <= write_target_row_max:
                # 結合対象セルかどうかを判定する
                is_target, merged_cells = is_target_merge(mergedcells_list, write_target_reference_column, write_row)
                if is_target:
                    target_cells_parent = merge_cells_parent(merged_cells, write_column)
                    opened_write_target.merge_cells(target_cells_parent)
                    print("[INFO] 結合対象セル: {colrow} を処理しました。".format(colrow=str(target_cells_parent)))
                # 行をインクリメントする
                write_row += 1
        
        # workbookのプロパティをリセットする(作成者:openpyxlを削除)
        opened_write_target.reset_workbook_properties()
        # 保存して終了する
        opened_write_target.close_workbook()
        opened_write_target.save_workbook()
        print("[INFO] すべての処理が正常に終了しました。")
        exitcode = 0
        sys.exit(exitcode)

    except Exception as e:
        # 異常があった場合はセーブしない
        print("[ERROR] 処理が異常終了しました。")
        print(str(e))
        print(str(traceback.format_exc()))
        exitcode = 1
        sys.exit(exitcode)

if __name__ == "__main__":
    main()
