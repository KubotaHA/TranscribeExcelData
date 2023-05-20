#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import traceback
import argparse

try:
    from transcribeexceldata import OpyxlWrapper
    from transcribeexceldata import check_if_file_exists
    from transcribeexceldata import get_yaml_data
except:
    print("[ERROR] transcribeexceldata.py が配置されていない可能性があります。")
    print(traceback.format_exc())

def main():
    try:
        # 引数パーサーの定義
        args_parser = argparse.ArgumentParser(
            # プログラム名
            prog=str(str(os.path.basename(__file__))),
            # Usage
            usage="{prog} -f <file path> -n <sheet number> [OPTIONS]\n  e.g. $ python3 {prog} -f target.xlsx -n 3".format(
                    prog=str(os.path.basename(__file__))),
            # 説明
            description="[Description: このプログラムは指定ファイル同士の指定セルの値を比較します。]",
            # -h/--helpヘルプオプションの有効化
            add_help=True
        )
        # 引数の定義
        args_parser.add_argument('-f', '--file-path', type=str, required=False, help="比較対象先ファイルパスを指定します。")
        args_parser.add_argument('-n', '--sheet-number', type=int, required=False, help="対象シート番号を指定します。(1 始まりです)")
        args_parser.add_argument('-y', '--yes-skip-warning', required=False, action='store_true', help="[OPTIONS] これを指定した場合は処理続行の警告メッセージをスキップします。")
        # 定義した引数の有効化
        args = args_parser.parse_args()
    except Exception as e:
        print("[ERROR: main] 処理が異常終了しました。")
        print(str(e))
        print("[ERROR: main] Traceback is as follows.....................")
        print(str(traceback.format_exc()))
        exitcode = 1
        sys.exit(exitcode)

    try:
        print("[INFO] 処理を開始します。")
        # 警告の表示
        print("[INFO] !!警告!! Excelファイル内の挿入図形やグラフが削除される可能性があります。")
        if args.yes_skip_warning != True:
            print("[WARN] --> 実行前にバックアップを推奨します。続行しますか？ (y/n)")
            continue_process = input('> ')
        else:
            continue_process = 'y'
        # 続行有無の判定
        if continue_process == 'y':
            pass
        else:
            print("[INFO] {continue_process} が指定されたため、処理を中断しました。".format(continue_process=continue_process))
            sys.exit(0)
        # fielddefinition.yml を読み込む
        fielddefinition = "fielddefinition.yml"
        if check_if_file_exists(fielddefinition) == False:
            raise Exception("[ERROR main] ファイル {file} が存在しません。".format(file=fielddefinition))
        yaml_data = get_yaml_data(fielddefinition)

        # 書き込み対象のExcelファイルを読み込む
        ## 読み取り先のExcelファイルについて引数指定がなかった場合は fielddefinition.yml から取得する
        if args.file_path == None:
            write_target_filename = yaml_data["checkcom_target"]["dst_file"]
        else:
            write_target_filename = args.file_path
        if check_if_file_exists(write_target_filename) == False:
            raise Exception("[ERROR main] ファイル {file} が存在しません。".format(file=write_target_filename))
        ## 対象シートについて引数指定がなかった場合は fielddefinition.yml から取得する
        if args.sheet_number == None:
            write_target_sheetnumber = yaml_data["checkcom_target"].get("src_sheet_number", None)
        else:
            write_target_sheetnumber = args.sheet_number
        write_target_sheetname = yaml_data["numberitems_target"].get("sheet_name", None)
        write_target_row = yaml_data["numberitems_target"]["row_definition"]
        write_target_row_max = yaml_data["numberitems_target"]["row_max"]
        write_target_column = yaml_data["numberitems_target"]["column_definition"]
        opened_write_target = OpyxlWrapper(write_target_filename)
        opened_write_target.load_workbook()
        opened_write_target.load_worksheet(write_target_sheetnumber, write_target_sheetname)

        # 列を軸としてループする
        ## 指定されたカラムの分だけループ処理する
        # 書き込み対象の行の開始番号を定義する
        write_row = write_target_row
        ## 項番を1から始める
        item_number = 1
        ## 行ループを開始する -> 処理する行の最大値までループする
        while write_row <= write_target_row_max:
            # 書き込み対象のカラムを指定する
            write_column = write_target_column
            # セルの値を取得する
            targetcell_value = opened_write_target.get_celldata(write_row, write_column)
            # セルの値が空でない場合のみセルへの書き込みを実行する
            if targetcell_value != None and targetcell_value != "":
                # セルの値(項番の番号)を書き込む
                opened_write_target.write_celldata(write_row, write_column, item_number)
                print("[INFO] Debug: CELL={col_row:4}: 書き込み項番: {item_num}".format(col_row=write_column+str(write_row), item_num=item_number))
                # 項番をインクリメントする
                item_number += 1
            # 行番号をインクリメントする
            write_row += 1
        
        # workbookのプロパティをリセットする(作成者:openpyxlを削除)
        opened_write_target.reset_workbook_properties()
        # 保存して終了する
        opened_write_target.save_workbook()
        opened_write_target.close_workbook()
        print("[INFO] すべての処理が正常に終了しました。")
        sys.exit(0)

    except Exception as e:
        opened_write_target.close_workbook()
        # 異常があった場合はセーブしない
        print("[ERROR] 処理が異常終了しました。")
        print(str(e))
        print("[ERROR: main] Traceback is as follows.....................")
        print(str(traceback.format_exc()))
        sys.exit(1)

if __name__ == "__main__":
    main()
