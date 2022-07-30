#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import traceback

try:
    from transcribeexceldata import OpyxlWrapper
    from transcribeexceldata import check_if_file_exists
    from transcribeexceldata import get_yaml_data
except:
    print("[ERROR] transcribeexceldata.py が配置されていない可能性があります。")
    print(traceback.format_exc())

def main():
    print("[INFO] 処理を開始します。")
    try:
        # fielddefinition.yml を読み込む
        fielddefinition = "fielddefinition.yml"
        if check_if_file_exists(fielddefinition) == False:
            raise Exception("[ERROR main] ファイル {file} が存在しません。".format(file=fielddefinition))
        yaml_data = get_yaml_data(fielddefinition)

        # 書き込み対象のExcelファイルを読み込む
        write_target_filename = yaml_data["numberitems_target"]["file_name"]
        if check_if_file_exists(write_target_filename) == False:
            raise Exception("[ERROR main] ファイル {file} が存在しません。".format(file=write_target_filename))
        write_target_sheetnumber = yaml_data["numberitems_target"].get("sheet_number", None)
        write_target_sheetname = yaml_data["numberitems_target"].get("sheet_name", None)
        write_target_row = yaml_data["numberitems_target"]["row_definition"]
        write_target_row_max = yaml_data["numberitems_target"]["row_max"]
        write_target_item_number_column = yaml_data["numberitems_target"]["item_number_column"]
        opened_write_target = OpyxlWrapper(write_target_filename)
        opened_write_target.load_workbook()
        opened_write_target.load_worksheet(write_target_sheetnumber, write_target_sheetname)

        # 書き込み対象の行の開始番号を定義する
        write_row = write_target_row
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
        print("[INFO] すべての処理が正常に終了しました。")
        exitcode = 0

    except Exception as e:
        # 異常があった場合はセーブしない
        print("[ERROR] 処理が異常終了しました。")
        print(str(e))
        print(str(traceback.format_exc()))
        exitcode = 1
    
    finally:
        opened_write_target.close_workbook()
        sys.exit(exitcode)


if __name__ == "__main__":
    main()
