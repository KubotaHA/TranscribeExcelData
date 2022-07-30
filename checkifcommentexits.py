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

def replace_blank(data):
    if isinstance(data, str) != True:
        raise Exception("[ERROR: replace_blank] data is not str. data: '{data}', type: {type}.".format(
            data=data, type=type(data)))
    return data.replace(" ", "").replace("　","")

def main():
    print("[INFO] 処理を開始します。")
    try:
        # fielddefinition.yml を読み込む
        fielddefinition = "fielddefinition.yml"
        if check_if_file_exists(fielddefinition) == False:
            raise Exception("[ERROR main] ファイル {file} が存在しません。".format(file=fielddefinition))
        yaml_data = get_yaml_data(fielddefinition)

        # 読み取り元のExcelファイルを読み込む
        src_filename = yaml_data["checkcom_target"]["src_file"]
        if check_if_file_exists(src_filename) == False:
            raise Exception("[ERROR main] ファイル {file} が存在しません。".format(file=src_filename))
        src_sheetnumber = yaml_data["checkcom_target"].get("src_sheet_number", None)
        src_row = yaml_data["checkcom_target"]["src_row_def"]
        src_row_max = yaml_data["checkcom_target"]["src_row_max"]
        src_column = yaml_data["checkcom_target"]["src_column_def"]
        opened_src_target = OpyxlWrapper(src_filename)
        opened_src_target.load_workbook()
        opened_src_target.load_worksheet(src_sheetnumber, None)

        # 読み取り先のExcelファイルを読み込む
        dst_filename = yaml_data["checkcom_target"]["dst_file"]
        if check_if_file_exists(dst_filename) == False:
            raise Exception("[ERROR main] ファイル {file} が存在しません。".format(file=dst_filename))
        dst_sheetnumber = yaml_data["checkcom_target"].get("src_sheet_number", None)
        dst_row = yaml_data["checkcom_target"]["dst_row_def"]
        dst_row_max = yaml_data["checkcom_target"]["dst_row_max"]
        dst_column = yaml_data["checkcom_target"]["dst_column_def"]
        opened_dst_target = OpyxlWrapper(dst_filename)
        opened_dst_target.load_workbook()
        opened_dst_target.load_worksheet(dst_sheetnumber, None)

        # 読み取り元のExcelファイルからデータを取得する
        src_comment_list = get_targetcelldata(opened_src_target, src_row, src_column, src_row_max)
        # 読み取り先のExcelファイルからデータを取得する
        dst_comment_list = get_targetcelldata(opened_dst_target, dst_row, dst_column, dst_row_max)

        # 読み取り先のデータと読み取り元データを比較して一致しているものを出力する
        result_list = []
        for dst in dst_comment_list:
            for src in src_comment_list:
                if replace_blank(src["value"]) == replace_blank(dst["value"]):
                    result_list.append(dst)

        # 結果の出力
        if len(result_list) > 0:
            print("[INFO] 一致したデータが見つかりました。以下の通りです。")
            for item in result_list:
                print(" * セル: {cell:5}-> '{data}'".format(
                    cell=item["cell"], data=item["value"]
                ))
        else:
            print("[INFO] 一致したデータはありませんでした。")

        # 終了
        print("[INFO] すべての処理が正常に終了しました。")
        exitcode = 0

    except Exception as e:
        # 異常があった場合はセーブしない
        print("[ERROR: main] 処理が異常終了しました。")
        print(str(e))
        print("[ERROR: main] Traceback is as follows.....................")
        print(str(traceback.format_exc()))
        exitcode = 1
    
    finally:
        opened_src_target.close_workbook()
        opened_dst_target.close_workbook()
        sys.exit(exitcode)



if __name__ == "__main__":
    main()
