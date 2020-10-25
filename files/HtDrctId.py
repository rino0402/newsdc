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
import unicodedata

def truncate(txt, num_bytes, encoding='shift_jis'):
    txt = txt.replace('\uff0d', '-')
    while len(txt.encode(encoding)) > num_bytes:
        txt = txt[:-1]

    return txt + ' '*(num_bytes - len(txt.encode(encoding)))

def main(r):
    print("main({})".format(r))
    print("pyodbc.connect({})".format(r["dns"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dns"])
    print(conn.getinfo(pyodbc.SQL_DATA_SOURCE_NAME))

    if r["table"] == "y_syuka":
        r["load"] = y_syuka(conn, r)
    else:
        r["load"] = hmtah015(conn, r)

    if r["uncommit"]:
        pass
    else:
        print("conn.commit()", end=".")
        conn.commit()
        print("ok")
  
    print("conn.close()", end=".")
    conn.close()
    print("ok")

    return r

def hmtah015(conn, r):
    where = ""
#    print(len(r["dt"]))
    if len(r["dt"]) == 1:
        where = " and h.SyukaDt = '{}'".format(r["dt"][0])
    elif len(r["dt"]) > 1:
        where = " and h.SyukaDt between '{}' and '{}'".format(r["dt"][0], r["dt"][1])
    sql = """
select distinct
 h.SyukaDt
,h.IDNo
,i.IDNo iIDNo
,h.ChoCode
,i.ChoCode iChoCode
,h.ChoName
,i.ChoName iChoName
,h.ChoZip
,i.ChoZip iChoZip
,h.ChoTel
,i.ChoTel iChoTel
,h.ChoAddress
,i.ChoAddress iChoAddress
,h.ChoMemo
,i.ChoMemo iChoMemo
,h.TMark
,i.TMark iTMark
from {} h
left outer join HtDrctId i on (h.IDNo=i.IDNo)
where (h.ChoCode <> i.ChoCode
or h.ChoName <> i.ChoName
or h.ChoZip <> i.ChoZip
or h.ChoTel <> i.ChoTel
or h.ChoAddress <> i.ChoAddress
or h.ChoMemo <> i.ChoMemo
or h.TMark <> i.TMark)
and h.ChoCode <> ''
{}
order by 1,2
""".format(r["table"], where)
    try:
        cursor = conn.cursor()
        cursor.execute(sql)
        columns = [column[0] for column in cursor.description]
        for row in cursor.fetchall():
            for i, col in enumerate(row):
                if isinstance(col, str):
                    row[i] = col.rstrip()
#            s = truncate(row.ChoName, 20)
#            print(s + ":" + str(len(s.encode('shift_jis'))))
            print("{:8.8} {} {:8.8} {} {:8.8}{:12.12} {} {} {}".format(
                 row.SyukaDt
                ,row.IDNo
                ,row.ChoCode
                ,truncate(row.ChoName, 30)
                ,row.ChoZip
                ,row.ChoTel
                ,truncate(row.ChoAddress, 30)
                ,truncate(row.ChoMemo, 30)
                ,row.TMark
                 ))
            print("{:8.8} {} {:8.8} {} {:8.8}{:12.12} {} {} {}".format(
                 ''
                ,row.iIDNo
                ,row.iChoCode
                ,truncate(row.iChoName, 30)
                ,row.iChoZip
                ,row.iChoTel
                ,truncate(row.iChoAddress, 30)
                ,truncate(row.iChoMemo, 30)
                ,row.iTMark
                 )
                  ,end=".")
            d = {}
            d["IDNo"] = row.IDNo
            d["ChoCode"] = row.ChoCode
            d["ChoName"] = row.ChoName
            d["ChoZip"] = row.ChoZip
            d["ChoTel"] = row.ChoTel
            d["ChoAddress"] = row.ChoAddress
            d["ChoMemo"] = row.ChoMemo
            d["TMark"] = row.TMark
            d["UpdID"] = r["table"]
            d["UpdTm"] = datetime.now().strftime("%Y%m%d%H%M%S")
            print("update:{}".format(update(conn, d)))
    except pyodbc.ProgrammingError as e:
        print(e)
        raise
    except pyodbc.Error as e:
        print(e)
        raise
    except:
        print('error')
        raise
    
def y_syuka(conn, r):
    sql = """
select
 y.KEY_ID_NO
,y.KEY_MUKE_CODE
,y.MUKE_NAME
,y.LK_MUKE_CODE
,i.ChoCode
,i.ChoName
,i.ChoZip
,i.ChoTel
,i.ChoAddress
,i.ChoMemo
,i.TMark
from y_syuka y
left outer join HtDrctId i on (y.KEY_ID_NO=i.IDNo)
where y.CHOKU_KBN = '1'
and i.IDNo is Null
order by 1
"""
    try:
        cursor = conn.cursor()
        cursor.execute(sql)
        columns = [column[0] for column in cursor.description]
        for row in cursor.fetchall():
            for i, col in enumerate(row):
                if isinstance(col, str):
                    row[i] = col.rstrip()
            print("{} {:8.8} {:8.8} {}".format(row.KEY_ID_NO,
                                 row.KEY_MUKE_CODE,
                                 row.LK_MUKE_CODE,
                                 row.MUKE_NAME
                                 )
                  ,end="."
                  )
            d = {}
            d["IDNo"] = row.KEY_ID_NO
            d["ChoCode"] = row.KEY_MUKE_CODE
            d["ChoName"] = row.MUKE_NAME
            d["EntID"] = "y_syuka"
            d["EntTm"] = datetime.now().strftime("%Y%m%d%H%M%S")
            print(insert(conn, d))
    except pyodbc.ProgrammingError as e:
        print(e)
        raise
    except pyodbc.Error as e:
        print(e)
        raise
    except:
        print('error')
        raise

def update(conn, data):
    sql = "update HtDrctId "
    sql += "set UpdID = '{0}'".format(data["UpdID"])
    for d in data:
        if d not in ["UpdID","IDNo"]:
            sql += ",{0} = '{1}'".format(d, data[d].replace("'","''"))
    sql += " where IDNo='{0}'".format(data["IDNo"])
#    print(sql, end=";")
    try:
        conn.execute(sql)
        ret = conn.execute("select @@rowcount").fetchone()[0]
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

def insert(conn, data):
    if data is None:
        return
    else:
        sql = "insert into HtDrctId (" + ",".join(map(str, data)) + ") values ("
        for d in data:
#           print("{0}:{1}".format(d, data[d]))
            if type(data[d]) == str:
                sql += "'{0}',".format(data[d])
            else:
                sql += "{0},".format(data[d])
            
        sql = sql[:-1] + ")"
    print(sql, end=".")
    try:
        conn.execute(sql)
        print("ok")
    except pyodbc.ProgrammingError as e:
        print(e)
        raise
    except pyodbc.Error as e:
        if 'SQLExecDirectW' in str(e.args):
            print("(W)")
#           raise
        else:
            print(e)
            raise
    except:
        print('error')
        raise

if __name__ == "__main__":
    dns = "newsdc"
    limit = 0
    if 'REQUEST_METHOD' in os.environ:
        pass
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: {0}".format(dns), default= dns, type= str)
        parser.add_argument("--table", help="y_syuka | HMTAH015".format(dns), default= "y_syuka", type= str)
        parser.add_argument("--dt", help="20201019", default= "", nargs="*", type= str)
        parser.add_argument("--limit", help="default: 0", default= limit, type= int)
        parser.add_argument("--uncommit", action='store_true')
#        args= parser.parse_args()
        r = main(vars(parser.parse_args()))
#        print(r)
