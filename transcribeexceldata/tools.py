#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import traceback
import yaml
import re

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

def readparent2writeparent(read_parent, read_row, write_row, write_column)->str:
    re_result = re.search(r'^([A-Z]+)([0-9]+)$', read_parent)
    if re_result == None:
        raise Exception("[ERROR readparent2writeparent] read_parent is incorrect format. -> " + str(read_parent))
    row_delta = write_row - read_row
    return write_column + str(int(re_result.group(2)) + row_delta)

def check_if_duplicated(columnlist):
    try:
        if columnlist == list:
            raise Exception("columnlist is not a list. type -> " + str(type(columnlist)))
        checkedlist = []
        for col in columnlist:
            if col in checkedlist:
                return False, col
            if re.search(r'^[A-Z]+$', col) == None:
                raise Exception("'{0}' is invalid format. e.g. 'A', 'AA'".format(str(col)))
            checkedlist.append(col)
        return True , None
    except Exception as e:
        raise Exception("[ERROR check_if_duplicated] Unexpected Error has been ocurred. -> " + str(e))

def transcribeexceldata():
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
        read_target_sheetnumber = yaml_data["read_target"].get("sheet_number", None)
        read_target_sheetname = yaml_data["read_target"].get("sheet_name", None)
        if read_target_sheetnumber == None and read_target_sheetname == None:
            raise Exception("[ERROR main] シート名かシート番号の指定が存在しません。")
        read_target_column = yaml_data["read_target"]["column_definition"]
        result_deplicated_bool, result_deplicated_col = check_if_duplicated(read_target_column)
        if result_deplicated_bool == False:
            raise Exception("[ERROR main] カラム名 '{col}' が重複しています。".format(col=result_deplicated_col))
        read_target_row = yaml_data["read_target"]["row_definition"]
        read_target_row_max = yaml_data["read_target"]["row_max"]
        opened_read_target = OpyxlWrapper(read_target_filename)
        opened_read_target.load_workbook()
        opened_read_target.load_worksheet(read_target_sheetnumber, read_target_sheetname)
        # 読み込み対象のExcelファイルから結合されたセルのリストを取得する
        mergedcells_list = opened_read_target.get_mergedcells_list()
        # 書き込み対象のExcelファイルを読み込む
        write_target_filename = yaml_data["write_target"]["file_name"]
        if check_if_file_exists(write_target_filename) == False:
            raise Exception("[ERROR main] ファイル {file} が存在しません。".format(file=write_target_filename))
        write_target_sheetnumber = yaml_data["write_target"].get("sheet_number", None)
        write_target_sheetname = yaml_data["write_target"].get("sheet_name", None)
        write_target_column = yaml_data["write_target"]["column_definition"]
        result_deplicated_bool, result_deplicated_col = check_if_duplicated(write_target_column)
        if result_deplicated_bool == False:
            raise Exception("[ERROR main] カラム名 '{col}' が重複しています。".format(col=result_deplicated_col))
        write_target_row = yaml_data["write_target"]["row_definition"]
        opened_write_target = OpyxlWrapper(write_target_filename)
        opened_write_target.load_workbook()
        opened_write_target.load_worksheet(write_target_sheetnumber, write_target_sheetname)
        # 転記先と転記元における対象カラムの数が合わない場合はエラー
        if len(read_target_column) != len(write_target_column):
            print("[ERROR main] 転記先と転記元における対象カラムの数が合いません。")

        # 列を軸としてループする
        ## 指定されたカラムの分だけループ処理する
        for column_num in range(len(read_target_column)):
            # 読み込み対象の行の開始番号を定義する
            read_row = read_target_row
            # 書き込み対象の行の開始番号を定義する
            write_row = write_target_row
            ## 行ループを開始する -> 処理する行の最大値までループする
            while read_row <= read_target_row_max:
                # 読み込み対象のセルを指定する
                read_column = read_target_column[column_num]
                # セルの値を取得する
                read_value = opened_read_target.get_celldata(read_row, read_column)
                # セルの色を取得する
                read_color, read_filltype = opened_read_target.get_cellcolor(read_row, read_column)
                # セルの書式設定を取得する
                read_alignment = opened_read_target.get_cellalignment(read_row, read_column)
                # 書き込み対象のカラムを指定する
                write_column = write_target_column[column_num]
                # "A1"の形式で読み込み対象のセルを定義する
                read_cell = read_column + str(read_row)
                # "A1"の形式で書き込み対象のセルを定義する
                write_cell = write_column + str(write_row)
                # セルのマージを実行する
                for merged_cells in mergedcells_list:
                    # merged_cellsリストの0番目がセル結合の開始セルとなるため
                    # それが読み込み対象のセルと一致している場合のみセル結合を行う
                    if merged_cells[0] == read_cell:
                        # merged_cellsリストの1番目がセル結合の終端セルとなるため
                        # その情報から書き込み対象セルの終端を求める
                        merge_cells_parent = write_cell + ":" + readparent2writeparent(merged_cells[1], read_target_row, write_target_row, write_column)
                        opened_write_target.merge_cells(merge_cells_parent)
                        print("[INFO] {needed_merge_cells} をセル結合しました。".format(needed_merge_cells=merge_cells_parent))
                        break
                # セルの値が空でない場合のみセルへの書き込みを実行する
                if read_value != None and read_value != "":
                    # セルの値を書き込む
                    opened_write_target.write_celldata(write_row, write_column, read_value)
                    # セルに色を適用する
                    opened_write_target.fill_cellcolor(write_row, write_column, read_color, read_filltype)
                    # セルの書式設定を適用する
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

        # workbookのプロパティをリセットする(作成者:openpyxlを削除)
        opened_write_target.reset_workbook_properties()
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
        opened_read_target.close_workbook()
        opened_write_target.close_workbook()
        sys.exit(exitcode)
