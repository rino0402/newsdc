# -*- coding: utf-8 -*-
import os
import sys
import io
import cgi
import cgitb
import numpy as np
import pandas as pd
import pyodbc
import json
from datetime import date, datetime, timedelta
from decimal import Decimal
#----------------------------------------------------------------
def main(r):
    print("main({})".format(r))
    print("pyodbc.connect({0})".format(r["dsn"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dsn"])
    print("ok")
    r["df"] = get_list(conn, r)
    print("conn.close()", end=".")
    conn.close()
    print("ok")
    return r
#----------------------------------------------------------------
def get_list(conn, r):
    sql = """
select
 left(JITU_DT,6) "処理年月"
,JITU_DT
,convert(right(JITU_DT,2),sql_integer)	"日"
,JITU_TM "処理時刻"
,p.TANTO_CODE + ' ' + ifnull(t.TANTO_NAME,'') "担当者"
,WEL_ID "WEL_ID"
,p.JGYOBU "事"
,p.MENU_NO + ' ' + ifnull(m.MENU_DSP,'') "MENU"
,case 
 when RIRK_ID in ('12','13','14','15','17','70')		then '入荷'
 when RIRK_ID = '80' and FROM_SOKO = '90'				then '入荷'
 when RIRK_ID in ('21','22','25','2A''42','43','4E')	then '出庫'
 when RIRK_ID in ('S2','S3')							then '出荷'
 end "kotei"
,RIRK_ID + ' ' + ifnull(y.YOIN_DNAME,'') "yoin"
,ID_NO
,HIN_GAI
,1 "Cnt"
,convert(SUMI_JITU_QTY,SQL_DECIMAL) + convert(MI_JITU_QTY,SQL_DECIMAL) 	"Qty"
,MUKE_CODE
,FROM_SOKO + FROM_RETU + FROM_REN + FROM_DAN "移動元"
,TO_SOKO + TO_RETU + TO_REN + TO_DAN "移動先"
,SHIJI_No
,convert(WORK_TM,SQL_DECIMAL) "作業時間(秒)"
,PRG_ID
From P_SAGYO_LOG p
left outer join YOIN y on (y.CODE_TYPE = left(RIRK_ID,1) and y.YOIN_CODE = right(RIRK_ID,1))
left outer join P_MENU as m on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)
left outer join TANTO as t on (p.TANTO_CODE = t.TANTO_CODE)
where p.JITU_DT between '{0}' and '{1}'
union
select
 Left(u.UKEIRE_DT,6)	"処理年月"
,u.UKEIRE_DT
,convert(right(u.UKEIRE_DT,2),sql_integer)	"日"
,''						"処理時刻"
,''						"担当者"
,''						"WEL_ID"
,''						"事"
,''						"MENU"
,'商品化'				"工程"
,'完了登録'				"要因"
,''						"ID-No."
,''						"品番"
,1 "Cnt"
,convert(u.UKEIRE_QTY,sql_decimal)		"数量"
,''						"向け先"
,''						"移動元"
,''						"移動先"
,u.SHIJI_NO				"指示書No."
,0						"作業時間(秒)"
,''						PRG_ID
from P_SUKEIRE u
where u.UKEIRE_DT between '{0}' and '{1}'
union
select
 Left(KENPIN_YMD,6)		"処理年月"
,KENPIN_YMD
,convert(right(KENPIN_YMD,2),sql_integer)	"日"
,KENPIN_HMS				"処理時刻"
,KENPIN_TANTO_CODE		"担当者"
,''						"WEL_ID"
,JGYOBU					"事"
,''						"MENU"
,'出荷伝票'				"工程"
,KEY_CYU_KBN + ' ' + CYU_KBN_NAME			"要因"
,KEY_ID_NO				"ID-No."
,KEY_HIN_NO				"品番"
,1 "Cnt"
,convert(JITU_SURYO,sql_decimal)		"数量"
,KEY_MUKE_CODE			"向け先"
,''						"移動元"
,''						"移動先"
,''						"指示書No."
,0						"作業時間(秒)"
,''						PRG_ID
from del_syuka
where KENPIN_YMD between '{0}' and '{1}'
union
select
 Left(KENPIN_YMD,6)		"処理年月"
,KENPIN_YMD
,convert(right(KENPIN_YMD,2),sql_integer)	"日"
,KENPIN_HMS				"処理時刻"
,KENPIN_TANTO_CODE		"担当者"
,''						"WEL_ID"
,JGYOBU					"事"
,''						"MENU"
,'出荷伝票'				"工程"
,KEY_CYU_KBN + ' ' + CYU_KBN_NAME			"要因"
,KEY_ID_NO				"ID-No."
,KEY_HIN_NO				"品番"
,1 "Cnt"
,convert(JITU_SURYO,sql_decimal)		"数量"
,KEY_MUKE_CODE			"向け先"
,''						"移動元"
,''						"移動先"
,''						"指示書No."
,0						"作業時間(秒)"
,''						PRG_ID
from y_syuka
where KENPIN_YMD between '{0}' and '{1}'

""".format(r["start"], r["end"])
    df = pd.read_sql(sql, conn)
    print(df)
    df2 = df.pivot_table(index=["kotei","yoin"], columns='JITU_DT', values=["Cnt","Qty"], aggfunc=np.sum)
    print(df2)
    return df2
#----------------------------------------------------------------
if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        cgitb.enable()
        form = cgi.FieldStorage()
        r = {}
        r["dsn"] = form.getvalue('dsn', 'newsdcn')
        r["start"] = form.getvalue('s', '')
        r["end"] = form.getvalue('e', '')
        sys.stdout = None
        r = main(r)
        sys.stdout = sys.__stdout__
        print('Content-Type:application/json; charset=UTF-8;\n')
        print(r["df"].to_json(orient= 'split', force_ascii= True))
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dsn", help="default: newsdcn", default="newsdcn", type=str)
        parser.add_argument("start", help="", nargs="?", default="", type=str)
        parser.add_argument("end", help="", nargs="?", default="", type=str)
        r = main(vars(parser.parse_args()))
        #print(r["df"].to_json(orient= 'split', force_ascii= True))
