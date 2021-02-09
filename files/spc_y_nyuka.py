# -*- coding: utf-8 -*-
import os
import sys
import io
import cgi
import re
import xlrd
import pandas as pd
import pyodbc
import json
from datetime import date, datetime, timedelta
from decimal import Decimal

def main(r):
    print("main({})".format(r))
    print("pyodbc.connect({0})".format(r["dns"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dns"])
    print("ok")

    r["y_nyuka"] = y_nyuka(conn, r)

    if r["commit"] == 1:
        print("conn.commit()", end=".")
        conn.commit()
        print("ok")

    print("conn.close()", end=".")
    conn.close()
    print("ok")
    
    return r

def y_nyuka(conn, r):
    sql = """
select
*
from spc_nyuka
where ware in (select distinct ware from spc_ware)
"""
    df = pd.read_sql(sql, conn)
    df["pn"] = df["pn"].str.rstrip()
    df["prod"] = df["prod"].str.rstrip()
    df["ware"] = df["ware"].str.rstrip()
    print(df)
    for i, row in df.iterrows():
        print(i, row)
        print(row["dt"])
        d = {}
        d["KAN_KBN"] = "9"
        d["DT_SYU"] = row["dtype"]
        d["JGYOBU"] = "Y"
        d["NAIGAI"] = "1"
        sql = "select max(TEXT_NO) from y_nyuka"
        sql += " where JGYOBU='Y' and NAIGAI='1'"
        sql += " and SYUKA_YMD = '{}'".format(row["dt"])
        text_no = conn.execute(sql).fetchone()[0]
        if text_no:
            text_no = int(text_no) + 1
        else:
            text_no = 1
        d["TEXT_NO"] = "{:09d}".format(text_no)  # Char(  9)
        d["JGYOBA"] = row.jcode     # Char(  8) R00013
        d["HIN_NO"] = row["pn"]
        #1234567890
        #Z211130010
        d["DEN_NO"] = row["id_no"]     # Char( 10)
        d["SURYO"] = "{}".format(int(row["qty"]))
        d["SYUKO_YMD"] = row["dt"]
        d["SYUKA_YMD"] = row["dt"]
        d["KAN_DT"] = row["dt"]
        d["HIN_NAI"] = row["prod"]
        d["LIST_OUT_END_F"] = "0"
        d["LIST_NYU_KANRI_F"] = "8"
        d["LIST_NYU_CHECK_F"] = "0"
        d["INS_TANTO"] = "SPAIS"        # Char(  5)
        d["Ins_DateTime"] = "{:%Y%m%d%H%M%S}".format(datetime.now())
        d["MOTO_PROG_ID"] = ""          # Char(  8)
        d["MUKE_CODE"] = row["ware"]    # Char(  8) 01A-11
        ret = insert(conn, d)

    return df

def insert(conn, data):
    sql = "insert into y_nyuka ("
    c = ""
    for d in data:
        if data[d]:
            sql += c + d
            c = ","
    sql += ") values ("
    for d in data:
        if data[d]:
            if type(data[d]) == str:
                sql += "'{0}',".format(data[d].replace("'","''"))
            else:
                sql += "{0},".format(data[d])
    sql = sql[:-1] + ")"
    print(sql, end=";")
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
        raise
    except:
        print('error')
        raise

if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        pass
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--commit", help="default: 0", default=0, type=int)
        args= parser.parse_args()
        r = main(vars(parser.parse_args()))
        import pprint
        pprint.pprint(r)
