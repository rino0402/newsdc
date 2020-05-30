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
#-------------------------------------------------------------
def main(r):
    print("main({})".format(r))
    log = os.path.dirname(os.path.abspath(__file__)) + '\\AcZaiko.log'
    try:
        r["log"] = open(log, mode='r').readlines()[-1]
        print(r["log"])
    except:
        r["log"] = ""
    if r["filename"] == "log":
        return r

    print("pyodbc.connect({0})".format(r["dns"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dns"])
    print("ok")

    if r["filename"] == "":
        r["list"] = list1(conn, r)
    else:
        r["load"] = load(conn, r)
        if len(r["load"]) > 0:
            print("commit()", end=".")
            conn.commit()
            print("ok")
        try:
            r["mtime"] = datetime.fromtimestamp(os.stat(r["filename"]).st_mtime).strftime('%Y/%m/%d %H:%M:%S')
            r["log"] = "{}\t{}件\t{}\n".format(os.path.basename(r["filename"]),len(r["load"]), r["mtime"])
        except:
            r["log"] = "{}\t{}件\n".format(r["name"],len(r["load"]))
        open(log, mode='a').write(r["log"])
    print("conn.close()", end=".")
    conn.close()
    print("ok")

    return r
#-------------------------------------------------------------
def list1(conn, r):
    print("list1()")
    sql = "select {}".format(' top {}'.format(r["limit"]) if r["limit"] > 0 else '')
    sql += " * from AcZaiko order by 1"
    data = []
    print(sql)
    try:
        cursor = conn.cursor()
        cursor.execute(sql)
        columns = [column[0] for column in cursor.description]
        for row in cursor.fetchall():
            for i, col in enumerate(row):
                if isinstance(col, str):
                    row[i] = col.rstrip()
                elif col is None:
                    row[i] = ""
            print(row)
            d = dict(zip(columns, row))
            data.append(d)
    except pyodbc.ProgrammingError as e:
        print(e)
        raise
    except pyodbc.Error as e:
        print(e)
        raise
    except:
        print('error')
        raise
    return data
#-------------------------------------------------------------
def load(conn, r):
    print("load()")
    print("pd.read_excel({0})".format(r["filename"]), end=".")
    df = pd.read_excel(r["filename"], dtype= 'object') #sheet_name= '進捗', 
    print("ok")
#    print(df)
    df = df.fillna("")
    df.rename(columns=lambda s: s.replace("\n",""), inplace=True)
    df.rename(columns=lambda s: s.replace("納期回答日　納期回答年月日","納期回答年月日"), inplace=True)
    print(df)
    print("delete from AcZan", end=".")
    conn.execute("delete from AcZaiko")
    print(conn.execute("select @@rowcount").fetchone()[0])
    ret = []
    for i, row in df.iterrows():
        if r["limit"] > 0 and i > r["limit"]:
            break
        for idx, c in enumerate(row):
            if isinstance(c, datetime):
                row[idx] = pd.to_datetime(c).strftime("%Y-%m-%d")
            elif isinstance(c, date):
                row[idx] = pd.to_datetime(c).strftime("%Y-%m-%d")
#        print(i,row)
        data = {}
        data["Pn"] = row["品目番号"]
        data["Syushi"] = "NAR"
        data["Qty"] = row["センター倉庫"]
        print("{0}:{1}".format(i, data))
        data["instert"] = insert(conn, data)
        ret.append(data)
    return ret
#-------------------------------------------------------------
def insert(conn, data):
    sql = "insert into AcZaiko (" + ",".join(map(str, data)) + ") values ("
    for d in data:
        if type(data[d]) == str:
            sql += "'{0}',".format(data[d].replace("'","''"))
        else:
            sql += "{0},".format(data[d])
    sql = sql[:-1] + ")"
    print(sql)
    try:
        conn.execute(sql)
        return conn.execute("select @@rowcount").fetchone()[0]
    except pyodbc.ProgrammingError as e:
        print(e)
        raise
    except pyodbc.Error as e:
        if 'SQLExecDirectW' in str(e.args):
            print("(W)")
            return 0
        else:
            print(e)
            raise
    except:
        print('error')
        raise
#----------------------------------------------------------------
def decimal_default_proc(obj):
    if isinstance(obj, Decimal):
        return float(obj)
    raise TypeError
#----------------------------------------------------------------
def print_json(results):
    print('Content-Type:application/json; charset=UTF-8;\n')
    print(json.dumps(results,
                     default=decimal_default_proc,
                     ensure_ascii=True,
                     indent=4,
                     sort_keys=False,
                     separators=(',', ': ')))
#----------------------------------------------------------------
#-------------------------------------------------------------
if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        form = cgi.FieldStorage()
        filename = ''
        if 'upload' in form:
            fileitem = form['upload']
            if fileitem.filename:
                fn = os.path.basename(fileitem.filename)
                open('files/upload/' + fn, 'wb').write(fileitem.file.read())
                filename = os.path.dirname(os.path.abspath(__file__)) + '\\upload\\' + fn
        else:
            filename = form.getvalue('file', '')
        r = {}
        r["dns"] = form.getvalue('dns', 'newsdc')
        r["filename"] = filename
        r["name"] = form.getvalue('name', '')
        r["limit"] = int(form.getvalue('limit', 0))
        sys.stdout = None
        r = main(r)
        sys.stdout = sys.__stdout__
        print_json(r)
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("filename", help="2.xlsx", nargs="?", default="", type=str)
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--limit", help="default: 0", default=0, type=int)
        r = main(vars(parser.parse_args()))
        print(r)
