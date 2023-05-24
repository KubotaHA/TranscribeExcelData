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

def get_targetcelldata(opened_target, row, col, row_max):
    # 読み取り元のExcelファイルからデータを取得する
    ## 空白セルの数をカウントするための変数
    null_count = 0
    ## 読み取りデータのリスト
    read_comment_list = []
    ## 行ループを開始する -> 処理する行の最大値までループする
    while row <= row_max:
        # セルの値を取得する
        targetcell_value = opened_target.get_celldata(row, col)
        # セルの値が空でない場合のみリストに追加する(半角全角スペースも除外)
        if targetcell_value != None and replace_blank(targetcell_value) != "":
            # セルの値をリストに追加する
            read_comment_list.append({
                "cell": col + str(row),
                "value": targetcell_value
                })
        # セルの値が空の場合はカウントをインクリメントする
        else:
            null_count += 1
        # 行番号をインクリメントする
        row += 1
    return read_comment_list

# データの中から半角/全角スペースを取り除く
def replace_blank(data):
    if isinstance(data, str) != True:
        str_data = str(data)
    else:
        str_data = data
    return str_data.replace(" ", "").replace("　","")

def main():
    try:
        # 引数パーサーの定義
        args_parser = argparse.ArgumentParser(
            # プログラム名
            prog=str(str(os.path.basename(__file__))),
            # Usage
            usage="{prog} -f <file path> -n <sheet number> \n  e.g. $ python3 {prog} -f target.xlsx -n 3".format(
                    prog=str(os.path.basename(__file__))),
            # 説明
            description="[Description: このプログラムは指定ファイル同士の指定セルの値を比較します。]",
            # -h/--helpヘルプオプションの有効化
            add_help=True
        )
        # 引数の定義
        args_parser.add_argument('-f', '--file-path', type=str, required=False, help="比較対象先ファイルパスを指定します。")
        args_parser.add_argument('-n', '--sheet-number', type=int, required=False, help="対象シート番号を指定します。(1 始まりです)")
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
        # 処理の開始
        print("[INFO] 処理を開始します...")
        print("[INFO]")
        # fielddefinition.yml を読み込む
        fielddefinition = "fielddefinition.yml"
        if check_if_file_exists(fielddefinition) == False:
            raise Exception("[ERROR main] ファイル '{file}' が存在しません。".format(file=fielddefinition))
        yaml_data = get_yaml_data(fielddefinition)

        # 読み取り元のExcelファイルを読み込む
        src_filename = yaml_data["checkcom_target"]["src_file"]
        if check_if_file_exists(src_filename) == False:
            raise Exception("[ERROR main] ファイル '{file}' が存在しません。".format(file=src_filename))
        src_sheetnumber = yaml_data["checkcom_target"].get("src_sheet_number", None)
        src_row = yaml_data["checkcom_target"]["src_row_def"]
        src_row_max = yaml_data["checkcom_target"]["src_row_max"]
        src_column = yaml_data["checkcom_target"]["src_column_def"]
        opened_src_target = OpyxlWrapper(src_filename)
        opened_src_target.load_workbook()
        opened_src_target.load_worksheet(src_sheetnumber, None)

        # 読み取り先のExcelファイルを読み込む
        ## 読み取り先のExcelファイルについて引数指定がなかった場合は fielddefinition.yml から取得する
        if args.file_path == None:
            dst_filename = yaml_data["checkcom_target"]["dst_file"]
        else:
            dst_filename = args.file_path
        if check_if_file_exists(dst_filename) == False:
            raise Exception("[ERROR main] ファイル '{file}' が存在しません。".format(file=dst_filename))
        ## 対象シートについて引数指定がなかった場合は fielddefinition.yml から取得する
        if args.sheet_number == None:
            dst_sheetnumber = yaml_data["checkcom_target"].get("dst_sheet_number", None)
        else:
            dst_sheetnumber = args.sheet_number
        dst_row = yaml_data["checkcom_target"]["dst_row_def"]
        dst_row_max = yaml_data["checkcom_target"]["dst_row_max"]
        dst_column = yaml_data["checkcom_target"]["dst_column_def"]
        opened_dst_target = OpyxlWrapper(dst_filename)
        opened_dst_target.load_workbook()
        opened_dst_target.load_worksheet(dst_sheetnumber, None)
        # ロードしたシート情報の表示
        print("[INFO] -> シート情報= NUM: {number}, NAME: '{name}'".format(
            number=str(dst_sheetnumber), name=opened_dst_target.worksheet_title))
        # 読み取り元のExcelファイルからデータを取得する
        src_comment_list = get_targetcelldata(opened_src_target, src_row, src_column, src_row_max)
        # 読み取り先のExcelファイルからデータを取得する
        dst_comment_list = get_targetcelldata(opened_dst_target, dst_row, dst_column, dst_row_max)
        # 読み取り先のデータと読み取り元データを比較して一致しているものを出力する
        result_list = []
        for dst in dst_comment_list:
            print("[INFO] Debug: CELL={cell}: '{value}'".format(cell=str(dst["cell"]), value=str(dst["value"])))
            for src in src_comment_list:
                if replace_blank(src["value"]) == replace_blank(dst["value"]):
                    result_list.append(dst)
        # 結果の出力
        if len(result_list) > 0:
            print("[INFO]")
            print("[INFO] !!! 一致したデータが見つかりました !!! 結果は以下の通りです。")
            for item in result_list:
                print(" * セル: {cell:5}-> '{data}'".format(
                    cell=item["cell"], data=item["value"]
                ))
        else:
            print("[INFO]")
            print("[INFO] 一致したデータはありませんでした。")
        # 終了
        print("[INFO] すべての処理が正常に終了しました。")
        opened_src_target.close_workbook()
        opened_dst_target.close_workbook()
        sys.exit(0)

    except Exception as e:
        print("[ERROR: main] 処理が異常終了しました。")
        print(str(e))
        print("[ERROR: main] Traceback is as follows.....................")
        print(str(traceback.format_exc()))
        if opened_src_target.workbook != None:
            opened_src_target.close_workbook()
        if opened_dst_target.workbook != None:
            opened_dst_target.close_workbook()
        sys.exit(1)

if __name__ == "__main__":
    main()
