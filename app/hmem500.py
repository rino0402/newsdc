# -*- coding: utf-8 -*-
import os
import sys
import socket
import io
import cgi
import cgitb
import pandas as pd
import pyodbc
import json
from datetime import date, datetime, timedelta
from decimal import Decimal
import traceback
import re

def main(r):
    print("main({})".format(r))
    print("pyodbc.connect({0})".format(r["dsn"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dsn"])
    print("ok")

    if r["item"]:
        r = item(conn, r)
    elif r["y_nyuka"]:
        r = y_nyuka(conn, r)
    elif r["zaiko"]:
        r = zaiko(conn, r)
    elif r["y_syuka"]:
        r = y_syuka(conn, r)
    else:
        if r["filename"]:
            r = load(conn, r)
        else:
            r = list0(conn, r)

    if r["commit"] == "1":
        print("conn.commit()", end=".")
        conn.commit()
        print("ok")

    print("conn.close()", end=".")
    conn.close()
    print("ok")
    
    return r

def list0(conn, r):
    sql = """
select top 10
 Filename
,count(*) cnt
,sum(if(SyushiCd in ('SJ','SA'),1,0)) sjsa
from hmem500
group by
 Filename
order by
 Filename desc
"""
    print(sql)
    df = pd.read_sql(sql, conn)
    print(df)
    return r
    
def zaiko(conn, r):
    sql = """
select
 z.YUKO_Z_QTY
,z.ID_NO2 zID_NO2
,y.*
,h.Filename
from hmem500 h
inner join y_nyuka y
on (h.JGYOBU = y.JGYOBU and h.DenDt = y.SYUKA_YMD and (h.SyoriMD + h.Bin + h.SeqNo) = y.TEXT_NO)
left outer join zaiko z
on ((z.Soko_No + z.Retu + z.Ren + z.Dan) = y.NYUKO_TANABAN
    and z.JGYOBU = y.JGYOBU and z.NAIGAI = y.NAIGAI and z.HIN_GAI = y.HIN_NO
    and z.NYUKA_DT = y.SYUKA_YMD)
where h.Filename like '{}'
and y.KAN_KBN = '0'
""".format(os.path.basename(r["filename"]))
    print(sql)
    df = pd.read_sql(sql, conn)
    print(df)
    for i, row in df.iterrows():
        print(i, row)
        for key in row.keys():
            print(key, row[key])
        for index, value in enumerate(row):
            if isinstance(value, str):
                row[index] = value.rstrip()
        # 在庫
        d = {}
        d["Soko_No"] = row.NYUKO_TANABAN[0:2]
        d["Retu"] = row.NYUKO_TANABAN[2:4]
        d["Ren"] = row.NYUKO_TANABAN[4:6]
        d["Dan"] = row.NYUKO_TANABAN[6:8]
        d["JGYOBU"] = row.JGYOBU
        d["NAIGAI"] = row.NAIGAI
        d["HIN_GAI"] = row.HIN_NO
        d["GOODS_ON"] = "1" #0:商品化済 1:未商品
        d["NYUKA_DT"] = row.SYUKA_YMD
        d["HIN_NAI"] = row.HIN_NAI
        d["LOCK_F"] = "0"
        d["GENSANKOKU"] = row.GENSANKOKU
        d["SHIIRE_WORK_CENTER"] = row.SHIIRE_WORK_CENTER
        d["ID_NO2"] = row.ID_NO
        d["YOSAN_FROM"] = row.YOSAN_FROM
        d["YOSAN_TO"] = row.H_SOKO  #YOSAN_TO
        if row.YUKO_Z_QTY:
            #if row.ID_NO != row.zID_NO2:
            d["YUKO_Z_QTY"] = "{}".format(int(row.YUKO_Z_QTY) + int(row.SURYO))
            update(conn, "zaiko", d)
        else:
            d["YUKO_Z_QTY"] = row.SURYO
            if insert(conn, "zaiko", d) == 0:
                sql = "update zaiko set YUKO_Z_QTY = convert(YUKO_Z_QTY,sql_decimal) + {}".format(row.SURYO)
                sql += " ,ID_NO2 = '+' + ID_NO2"
                sql += " where Soko_No='{}'".format(d["Soko_No"])
                sql += " and Retu='{}'".format(d["Retu"])
                sql += " and Ren='{}'".format(d["Ren"])
                sql += " and Dan='{}'".format(d["Dan"])
                sql += " and JGYOBU='{}'".format(d["JGYOBU"])
                sql += " and NAIGAI='{}'".format(d["NAIGAI"])
                sql += " and HIN_GAI='{}'".format(d["HIN_GAI"])
                sql += " and GOODS_ON='{}'".format(d["GOODS_ON"])
                sql += " and NYUKA_DT='{}'".format(d["NYUKA_DT"])
                conn.execute(sql)

        # 移動履歴
        ido = {}
        ido["JITU_DT"] = row.Ins_DateTime[:8]
        ido["JITU_TM"] = row.Ins_DateTime[8:]
        ido["JGYOBU"] = row.JGYOBU
        ido["NAIGAI"] = row.NAIGAI
        ido["HIN_GAI"] = row.HIN_NO
        ido["RIRK_ID"] = "10"
        ido["SUMI_JITU_QTY"] = "00000000"
        ido["MI_JITU_QTY"] = row.SURYO
        ido["TO_SOKO"] = row.NYUKO_TANABAN[0:2]
        ido["TO_RETU"] = row.NYUKO_TANABAN[2:4]
        ido["TO_REN"] = row.NYUKO_TANABAN[4:6]
        ido["TO_DAN"] = row.NYUKO_TANABAN[6:8]
        ido["DEN_DT"] = row.SYUKO_YMD
        ido["DEN_NO"] = row.DEN_NO
        ido["PRG_ID"] = "HMEM500"
        ido["HIN_NAI"] = row.HIN_NAI
        ido["NYUKA_DT"] = row.SYUKA_YMD
        ido["NYUKO_DT"] = row.SYUKO_YMD
        ido["WEL_ID"] = socket.gethostname()
        ido["RIRK_NAME"] = "通常入荷"
        ido["HIN_NAME"] = row.HIN_NAME
        sql = """
select
 sum(if(GOODS_ON = '0',convert(YUKO_Z_QTY,SQL_DECIMAL),0)) qty0
,sum(if(GOODS_ON <> '0',convert(YUKO_Z_QTY,SQL_DECIMAL),0)) qty1
from zaiko
where jgyobu = '{}'
and naigai = '{}'
and hin_gai = '{}'
""".format(row.JGYOBU, row.NAIGAI, row.HIN_NO)
        z = conn.execute(sql).fetchone()
        ido["SUMI_HIN_Zaiko_Qty"] = "{}".format(int(z.qty0))
        ido["MI_HIN_Zaiko_Qty"] = "{}".format(int(z.qty1))
        ido["SUMI_FROM_TANA_Zaiko"] = "0"
        ido["SUMI_TO_TANA_Zaiko"] = "0"
        ido["MI_FROM_TANA_Zaiko"] = "0"
        ido["MI_TO_TANA_Zaiko"] = "0"
        ido["MEMO"] = ""
        ido["MUKE_DNAME"] = row.YOSAN_FROM
        ido["TANTO_NAME"] = row.Filename
        ido["SUM_KBN"] = "1"
        ido["ID_NO"] = row.ID_NO
        ido["Ins_DateTime"] = datetime.now().strftime("%Y%m%d%H%M%S")
        ido["JITU_DT"] = ido["Ins_DateTime"][:8]
        ido["JITU_TM"] = ido["Ins_DateTime"][8:]
        ido["GENSANKOKU"] = row.GENSANKOKU
        ido["SHIIRE_WORK_CENTER"] = row.SHIIRE_WORK_CENTER
        ido["ID_NO2"] = row.ID_NO2
        ido["YOSAN_FROM"] = row.YOSAN_FROM
        ido["YOSAN_TO"] = row.YOSAN_TO
        insert(conn, "idoreki", ido)

        if int(row.MAEGARI_SURYO or 0) > 0:
            # 在庫 前借相殺
            d["YUKO_Z_QTY"] = "{}".format(int(d["YUKO_Z_QTY"]) - int(row.MAEGARI_SURYO))
            if int(d["YUKO_Z_QTY"]) > 0:
                update(conn, "zaiko", d)
            else:
                sql = "delete from zaiko"
                sql += " where Soko_No='{}'".format(d["Soko_No"])
                sql += " and Retu='{}'".format(d["Retu"])
                sql += " and Ren='{}'".format(d["Ren"])
                sql += " and Dan='{}'".format(d["Dan"])
                sql += " and JGYOBU='{}'".format(d["JGYOBU"])
                sql += " and NAIGAI='{}'".format(d["NAIGAI"])
                sql += " and HIN_GAI='{}'".format(d["HIN_GAI"])
                sql += " and GOODS_ON='{}'".format(d["GOODS_ON"])
                sql += " and NYUKA_DT='{}'".format(d["NYUKA_DT"])
                conn.execute(sql)
            # 移動履歴 前借相殺
            ido["SUMI_JITU_QTY"] = "{}".format(int(row.MAEGARI_SURYO))
            ido["SUMI_HIN_Zaiko_Qty"] = "{}".format(int(ido["SUMI_HIN_Zaiko_Qty"])-int(row.MAEGARI_SURYO))
            ido["FROM_SOKO"] = ido["TO_SOKO"]
            ido["FROM_RETU"] = ido["TO_RETU"]
            ido["FROM_REN"] = ido["TO_REN"]
            ido["FROM_DAN"] = ido["TO_DAN"]
            ido["TO_SOKO"] = ""
            ido["TO_RETU"] = ""
            ido["TO_REN"] = ""
            ido["TO_DAN"] = ""
            ido["RIRK_ID"] = "20"
            ido["RIRK_NAME"] = "前借相殺"
            ido["SUM_KBN"] = "3"
            ido["Ins_DateTime"] = datetime.now().strftime("%Y%m%d%H%M%S")
            ido["JITU_DT"] = ido["Ins_DateTime"][:8]
            ido["JITU_TM"] = ido["Ins_DateTime"][8:]
            insert(conn, "idoreki", ido)

        # y_nyuka 完了セット
        y = {}
        y["JGYOBU"] = row.JGYOBU
        y["SYUKA_YMD"] = row.SYUKA_YMD
        y["TEXT_NO"] = row.TEXT_NO
        y["KAN_KBN"] = "9"
        y["UPD_TANTO"] = "ZAIKO"
        y["Upd_DateTime"] = datetime.now().strftime("%Y%m%d%H%M%S")
        update(conn, "y_nyuka", y)

    return r

def y_syuka(conn, r):
    sql = """
select
 y.KEY_ID_NO
,h.*
from hmem500 h
left outer join y_syuka y
on (h.JGYOBU = y.JGYOBU and h.ID_NO = y.KEY_ID_NO)
where h.Filename like '{}'
and h.SyushiCd in ('SJ')
and h.NyukoCd <> ''
and h.IoKbn = '2'
""".format(os.path.basename(r["filename"]))
    print(sql)
    df = pd.read_sql(sql, conn)
    print(df)
    for i, row in df.iterrows():
        print(i, row)
        for index, value in enumerate(row):
            if isinstance(value, str):
                row[index] = value.rstrip()
        d = {}
        d["DT_SYU"] = "0"
        d["JGYOBU"] = row.JGyobu
        d["KEY_CYU_KBN"] = "2"
        d["KEY_ID_NO"] = row.ID_NO
        d["NAIGAI"] = "1"
        d["KEY_HIN_NO"] = row.Pn
        d["KEY_MUKE_CODE"] = row.SyushiCd
        d["KEY_SYUKA_YMD"] = row.DenDt
        d["ID_NO"] = row.ID_NO
        d["HIN_NO"] = row.Pn
        d["DEN_NO"] = row.DenNo
        d["SURYO"] = row.Qty.strip()
        d["MUKE_CODE"] = row.SyushiCd
        d["SYUKO_SYUSI"] = row.SyushiCd
        d["SYUKO_YMD"] = row.DenDt
        d["SYUKA_YMD"] = row.DenDt
        d["CYU_KBN"] = d["KEY_CYU_KBN"]
        d["HIN_NAME"] = row.PName
        d["HIN_NAI"] = row.PnNai
        if row.KEY_ID_NO:
            d["UPD_NOW"] = re.findall(r'\d+-\d+', row.Filename)[0].replace('-','')
            update(conn, "y_syuka", d)
        else:
            d["KAN_KBN"] = "0"
            d["JITU_SURYO"] = "0000000"
            d["INS_NOW"] = re.findall(r'\d+-\d+', row.Filename)[0].replace('-','')
            insert(conn, "y_syuka", d)
        
    return r

def y_nyuka(conn, r):
    sql = """
select
 y.TEXT_NO
,h.*
from hmem500 h
left outer join y_nyuka y
on (h.JGYOBU = y.JGYOBU and h.DenDt = y.SYUKA_YMD and (h.SyoriMD + h.Bin + h.SeqNo) = y.TEXT_NO)
where h.Filename like '{}'
and h.IoKbn = '1'
and ((h.JGYOBU = 'A' and h.SyushiCd in ('SJ'))
  or (h.JGYOBU = 'N')
  )
""".format(os.path.basename(r["filename"]))
    print(sql)
    df = pd.read_sql(sql, conn)
    print(df)
    for i, row in df.iterrows():
        for index, value in enumerate(row):
            if isinstance(value, str):
                row[index] = value.rstrip()
        if row.JGyobu == "N":
            if row.SyukoCd[2:4] == row.SyushiCd:
                print("pass:", i, row)
                continue
            row.JGyobu = "N"
        d = {}
        d["DT_SYU"] = row.IoKbn
        d["JGYOBU"] = row.JGyobu
        d["NAIGAI"] = "1"
        d["TEXT_NO"] = row.SyoriMD + row.Bin + row.SeqNo
        d["ID_NO"] = row.ID_NO
        d["HIN_NO"] = row.Pn
        d["DEN_NO"] = row.DenNo
        d["SURYO"] = row.Qty.strip()
        d["SYUKO_YMD"] = row.DenDt
        d["SYUKA_YMD"] = row.DenDt
        d["HIN_NAME"] = row.PName
        d["NOUKI_YMD"] = row.SHITEI_NOUKI_YMD
        d["KAN_DT"] = row.DenDt
        d["YOSAN_FROM"] = row.SyukoCd
        d["YOSAN_TO"] = row.NyukoCd
        d["HTANABAN"] = row.Loc1
        d["HIN_NAI"] = row.PnNai
        d["H_SOKO"] = row.SyushiCd
        d["GENSANKOKU"] = row.GENSANKOKU
        d["GEN_GENSANKOKU"] = row.GEN_GENSANKOKU
        d["SHIIRE_WORK_CENTER"] = row.SHIIRE_WORK_CENTER
        d["KANKYO_KBN"] = row.KANKYO_KBN
        d["KANKYO_KBN_ST"] = row.KANKYO_KBN_ST
        d["KANKYO_KBN_SURYO"] = row.KANKYO_KBN_SURYO.strip()
        d["ID_NO2"] = row.ID_NO
        d["AITESAKI_CODE"] = row.AITESAKI_CODE
        d["JYUCHU_YMD"] = row.JYUCHU_YMD
        d["SHITEI_NOUKI_YMD"] = row.SHITEI_NOUKI_YMD
        d["LIST_OUT_END_F"] = "9"
        d["LIST_NYU_KANRI_F"] = "9"
        d["LIST_NYU_CHECK_F"] = "9"
        d["NYUKO_TANABAN"] = "90010101"
        if row.TEXT_NO:
            d["UPD_TANTO"] = "HMEM500"
            d["Upd_DateTime"] = re.findall(r'\d+-\d+', row.Filename)[0].replace('-','')
            update(conn, "y_nyuka", d)
        else:
            d["KAN_KBN"] = "0"
            sql = "select sum(convert(JITU_QTY,sql_decimal)) from j_nyuka"
            sql += " where JGYOBU='{}'".format(d["JGYOBU"])
            sql += " and NAIGAI='{}'".format(d["NAIGAI"])
            sql += " and HIN_GAI='{}'".format(d["HIN_NO"])
            jQty = conn.execute(sql).fetchone()[0]
            if jQty:
                d["MAEGARI_SURYO"] = "{}".format(min(int(row.Qty),jQty))
                sql = "delete from j_nyuka"
                sql += " where JGYOBU='{}'".format(d["JGYOBU"])
                sql += " and NAIGAI='{}'".format(d["NAIGAI"])
                sql += " and HIN_GAI='{}'".format(d["HIN_NO"])
                conn.execute(sql)
                jQty -= int(row.Qty)
                if jQty > 0:
                    j = {}
                    j["JGYOBU"] = d["JGYOBU"]
                    j["NAIGAI"] = d["NAIGAI"]
                    j["HIN_GAI"] = d["HIN_NO"]
                    j["JITU_QTY"] = "{}".format(int(jQty))
                    j["INS_DATE"] = datetime.now().strftime("%Y%m%d")
                    insert(conn, "j_nyuka", j)
                    
            d["INS_TANTO"] = "HMEM500"
            d["Ins_DateTime"] = re.findall(r'\d+-\d+', row.Filename)[0].replace('-','')
            insert(conn, "y_nyuka", d)

    return r


def item(conn, r):
    sql = """
select
 i.HIN_GAI
,h.*
from hmem500 h
left outer join item i
on (h.JGYOBU = i.JGYOBU and i.NAIGAI = '1' and h.Pn = i.HIN_GAI)
where i.HIN_GAI is null
and Filename like '{}'
and (h.JGYOBU = 'N' or (h.JGYOBU = 'A' and SyushiCd in ('SJ','SA')))
""".format(os.path.basename(r["filename"]))
    print(sql)
    df = pd.read_sql(sql, conn)
    print(df)
    for i, row in df.iterrows():
        print(i, row)
        for key in row.keys():
            print(key, row[key])
        for index, value in enumerate(row):
            if isinstance(value, str):
                row[index] = value.rstrip()
        d = {}
        d["JGYOBU"] = row.JGyobu
        d["NAIGAI"] = "1"
        d["HIN_GAI"] = row.Pn
        d["HIN_NAI"] = row.PnNai
        d["HIN_NAME"] = row.PName
        d["INS_TANTO"] = "HMEM500"
        d["Ins_DateTime"] = re.findall(r'\d+-\d+', row.Filename)[0].replace('-','')
        insert(conn, "item", d)

    return r

def update(conn, table, data):
    if table == "item":
        keys = ["JGYOBU","KEY_ID_NO"]
        where = " where JGYOBU='{0}'".format(data["JGYOBU"])
        where += " and KEY_ID_NO='{0}'".format(data["KEY_ID_NO"])
    elif table == "y_nyuka":
        keys = ["JGYOBU","SYUKA_YMD","TEXT_NO"]
        where = " where JGYOBU='{}'".format(data["JGYOBU"])
        where += " and SYUKA_YMD='{}'".format(data["SYUKA_YMD"])
        where += " and TEXT_NO='{}'".format(data["TEXT_NO"])
    elif table == "y_syuka":
        keys = ["JGYOBU","KEY_ID_NO"]
        where = " where JGYOBU='{}'".format(data["JGYOBU"])
        where += " and KEY_ID_NO='{}'".format(data["KEY_ID_NO"])
    elif table == "zaiko":
        keys = ["Soko_No","Retu","Ren","Dan","JGYOBU","NAIGAI","HIN_GAI","SYUKA_YMD"]
        where = " where Soko_No='{}'".format(data["Soko_No"])
        where += " and Retu='{}'".format(data["Retu"])
        where += " and Ren='{}'".format(data["Ren"])
        where += " and Dan='{}'".format(data["Dan"])
        where += " and JGYOBU='{}'".format(data["JGYOBU"])
        where += " and NAIGAI='{}'".format(data["NAIGAI"])
        where += " and HIN_GAI='{}'".format(data["HIN_GAI"])
        where += " and NYUKA_DT='{}'".format(data["NYUKA_DT"])
    else:
        return 0
    sql = "update {} ".format(table)
    delim = "set"
    for d in data:
        if d not in keys:
            sql += delim
            delim = ","
            sql += " {0} = '{1}'".format(d, data[d].replace("'","''"))
    sql += where
    print(sql, end=".")
    conn.execute(sql)
    ret = conn.execute("select @@rowcount").fetchone()[0]
    print(ret)
    return ret

def insert(conn, table, data):
    sql = "insert into " + table + " (" + ",".join(map(str, data)) + ") values ("
    for d in data:
        sql += "'{0}',".format(data[d])
    sql = sql[:-1] + ")"
    print(sql, end=".", flush=True)
    try:
        conn.execute(sql)
        ret = conn.execute("select @@rowcount").fetchone()[0]
        print(ret)
        return ret
    except pyodbc.ProgrammingError as e:
        print(e)
        raise
    except pyodbc.Error as e:
        print(e)
        if '(Btrieve Error 5) (-4994)' in str(e.args):
            return 0
        else:
            raise
    except:
        print('error')
        raise

import codecs
def load(conn, r):
    basename = os.path.basename(r["filename"])
    with open(r["filename"], mode='r', encoding='shift_jis') as f:
        sql = "delete from hmem500R where Filename = '{}'".format(basename)
        print(sql, end=" ; ")
        conn.execute(sql)
        print(conn.execute("select @@rowcount").fetchone()[0])
        for i, line in enumerate(f.readlines(), start=1):
            print(basename, i, line)
            sql = "insert into hmem500R (Filename,Row,RecBuff) values ('{}',{},'{}')".format(basename, i, line)
            conn.execute(sql)
            print("rowcount={}".format(conn.execute("select @@rowcount").fetchone()[0]))
    return r

if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        pass
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("filename", help="\\\\hs1\\gift\\recv\\hmem506szz.dat.20210216-132020.4732", nargs='?', default='', type=str)
        #parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--dsn", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--commit", help="0 1", default="1", type=str)
        parser.add_argument("--item", help="品目マスター登録", action="store_true")
        parser.add_argument("--y_nyuka", help="入荷予定登録", action="store_true")
        parser.add_argument("--zaiko", help="在庫データ登録", action="store_true")
        parser.add_argument("--y_syuka", help="出荷予定登録※子部品", action="store_true")
        r = main(vars(parser.parse_args()))
