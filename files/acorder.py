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
def main(dns, limit, filename):
    print("main({0}, {1}, {2})".format(dns, limit, filename))
    results = {}
    results["dns"] = dns
    results["limit"] = limit
    results["filename"] = filename
    results["table"] = 'AcOrder'
    log = os.path.dirname(os.path.abspath(__file__)) + '\\AcOrder.log'
    try:
        results["log"] = open(log, mode='r').readlines()[-1]
        print(results["log"])
    except:
        results["log"] = ""
    if filename == "log":
        return results

    print("pyodbc.connect({0})".format(dns), end=".")
    conn = pyodbc.connect('DSN=' + dns)
    print("ok")

    if filename == "":
        results["list"] = list1(conn, limit)
    else:
        results["load"] = load(conn, filename, limit)
        if len(results["load"]) > 0:
            print("commit()", end=".")
            conn.commit()
            print("ok")
        try:
            results["mtime"] = datetime.fromtimestamp(os.stat(filename).st_mtime).strftime('%Y/%m/%d %H:%M:%S')
            results["log"] = "{}\t{}件\t{}\n".format(os.path.basename(filename),len(results["load"]), results["mtime"])
        except:
            results["log"] = "{}\t{}件\n".format(results["filename"],len(results["load"]))
        open(log, mode='a').write(results["log"])
    print("conn.close()", end=".")
    conn.close()
    print("ok")

    return results
#-------------------------------------------------------------
def list1(conn, limit):
    print("list1()")
    sql = "select {}".format(' top {}'.format(limit) if limit > 0 else '')
    sql += " * from AcOrder order by 1"
    r = []
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
            r.append(d)
    except pyodbc.ProgrammingError as e:
        print(e)
        raise
    except pyodbc.Error as e:
        print(e)
        raise
    except:
        print('error')
        raise
    return r
#-------------------------------------------------------------
def load(conn, filename, limit):
    print("load({0}, {1})".format(filename, limit))
    print("pd.read_excel({0})".format(filename), end=".")
    df = pd.read_excel(filename, sheet_name= '進捗', dtype= 'object')
    print("ok")
#    print(df)
    r = []
    df = df.fillna("")
    df.rename(columns=lambda s: s.replace("\n",""), inplace=True)
    df.rename(columns=lambda s: s.replace("納期回答日　納期回答年月日","納期回答年月日"), inplace=True)
    print(df)
    print("delete from AcOrder", end=".")
    conn.execute("delete from AcOrder")
    print(conn.execute("select @@rowcount").fetchone()[0])
    for i, row in df.iterrows():
        if limit > 0 and i > limit:
            break
        for idx, c in enumerate(row):
            if isinstance(c, datetime):
                row[idx] = pd.to_datetime(c).strftime("%Y-%m-%d")
            elif isinstance(c, date):
                row[idx] = pd.to_datetime(c).strftime("%Y-%m-%d")
#        print(i,row)
        data = {}
        data["Row"] = i + 2
        data["IdNo"] = str(row["発注納入管理番号"])
        data["Pn"] = row["品目番号"]
        data["Qty"] = row["入出庫予定数"]
        if isinstance(data["Qty"], int):
            pass
        else:
            data["Qty"] = 0
        data["QtyS"] = row["正味入出庫予定数"]
        data["Noki"] = row["納期回答年月日"]
        data["Tanto"] = row["担当者名"]
        try:
            data["Biko1"] = str(row["備考欄１"])
        except:
            data["Biko1"] = str(row["備考欄"])

        print("{0}:{1}".format(row["商品化完了日"], type(row["商品化完了日"])))
        if isinstance(row["商品化完了日"], int):
            data["KanDt"] = (datetime(1899, 12, 30) + timedelta(days=row["商品化完了日"])).strftime("%Y-%m-%d")
            print("→{}".format(data["KanDt"]))
        else:        
            data["KanDt"] = row["商品化完了日"]
        data["KanQty"] = row["商品化完了数"]
        data["NaraDt"] = row["奈良納入日"]
        print("{0}:{1}".format(i, data))
        data["instert"] = insert(conn, data)
        r.append(data)
    return r
#-------------------------------------------------------------
def insert(conn, data):
    sql = "insert into AcOrder (" + ",".join(map(str, data)) + ") values ("
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
        dns = form.getvalue('dns', 'newsdc')
        limit = int(form.getvalue('limit', 0))
        sys.stdout = None
        r = main(dns, limit, filename)
        sys.stdout = sys.__stdout__
        print_json(r)
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("filename", help="発注残データ20180531.xlsx", nargs="?", default="", type=str)
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--limit", help="default: 0", default=0, type=int)
        args= parser.parse_args()
        r = main(args.dns, args.limit, args.filename)
