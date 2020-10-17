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
import pathlib

def main(dns, filename, limit):
    print("main({0}, {1}, {2})".format(dns, filename, limit))
    results = {}
    results["dns"] = dns
    results["limit"] = limit

    if pathlib.Path(filename).is_absolute():
        filepath = filename
    else:
        filepath = os.path.dirname(os.path.abspath(__file__)) + '\\' + filename
    results["filename"] = filename
    results["filepath"] = filepath
    results["mtime"] = datetime.fromtimestamp(os.stat(filepath).st_mtime)
    print("pyodbc.connect({0})".format(dns), end=".", flush=True)
    conn = pyodbc.connect('DSN=' + dns)
    print(conn.getinfo(pyodbc.SQL_DATA_SOURCE_NAME), flush=True)

    df = pd.read_excel(filepath, sheet_name="商品化予定")
    results["columns"] = df.to_dict(orient='split')['columns']
    print(results["columns"])

    results["data"] = load(conn, df, limit)

    print("conn.close()", end=".", flush=True)
    conn.close()
    print("ok", flush=True)

    return results

def load(conn, df, limit):
    print("load({0}, {1}, {2})".format(conn.getinfo(pyodbc.SQL_DATA_SOURCE_NAME), len(df), limit))
    df = df.fillna("")
    r = []
    for i, row in df.iterrows():
        print("{}".format(row))
        data = {}
        if isinstance(row["予定日"], date ):
            data["Dt"] = "{0}/{1}".format(row["予定日"].month, row["予定日"].day)
            data["Dt8"] = row["予定日"].strftime("%Y%m%d")
        else:
            data["Dt"] = row["予定日"]
            data["Dt8"] = row["予定日"]
        data["Pn"] = row["品番"]
        data["Qty"] = row["数量"]
        data["Sample"] = row[3] # row["見本"]
        data["Memo"] = row[4]   # row["備考"]
#        data["QtyE"] = row["完了数"]
        item = get_item(conn, data["Pn"], data["Dt8"])
        data["stts"] = ""
        data["rank"] = 0
        if item:
            data["HIN_NAME"] = item.HIN_NAME.rstrip()
            if data["HIN_NAME"] == "":
                data["HIN_NAME"] = item.L_HIN_NAME_E.rstrip()
            data["zQty"] = item.zQty if item.zQty != 0 else ""
            data["zQty92"] = item.zQty92 if item.zQty92 != 0 else ""
            data["zQtySumi"] = item.zQtySumi if item.zQtySumi != 0 else ""
            data["zQtyMi"] = item.zQtyMi if item.zQtyMi != 0 else ""
            data["sQty"] = item.sQty if item.sQty != 0 else ""
            data["uQty"] = item.uQty if item.uQty != 0 else ""
            if item.uQty > 0:
                data["stts"] = "uQty"
                data["rank"] = 1
            elif item.zQty92 >= data["Qty"]:
                data["stts"] = "zQty92"
            elif item.sQty >= data["Qty"]:
                data["stts"] = "sQty"
                data["rank"] = 1
#            elif item.zQtySumi >= data["Qty"]:
#                data["stts"] = "zQtyS"
            elif item.zQtyMi >= data["Qty"]:
                data["stts"] = "zQtyM"
        else:
            data["HIN_NAME"] = "*未登録*"
            data["zQty"] = ""
            data["zQty92"] = ""
            data["zQtySumi"] = ""
            data["zQtyMi"] = ""
            data["sQty"] = ""
            
        r.append(data)
    return sorted(r, key=lambda x:x["rank"])

def get_item(conn, pn, dt8):
    print("get_item({0}, {1})".format(conn.getinfo(pyodbc.SQL_DATA_SOURCE_NAME), pn))
    sql = """select
 i.HIN_NAME
,i.L_HIN_NAME_E
,ifnull(z.qty,0) zQty
,ifnull(z.qty92,0) zQty92
,ifnull(z.qtySumi,0) zQtySumi
,ifnull(z.qtyMi,0) zQtyMi
,ifnull(s.qty,0) sQty
,ifnull(u.qty,0) uQty
from item i
left outer join (
    select
    JGYOBU
    ,NAIGAI
    ,HIN_GAI
    ,sum(convert(YUKO_Z_QTY,sql_decimal)) qty
    ,sum(if(Soko_No = '92',convert(YUKO_Z_QTY,sql_decimal),0)) qty92
    ,sum(if(Soko_No <> '92' and GOODS_ON =  '0',convert(YUKO_Z_QTY,sql_decimal),0)) qtySumi
    ,sum(if(Soko_No <> '92' and GOODS_ON <> '0',convert(YUKO_Z_QTY,sql_decimal),0)) qtyMi
    from zaiko
    where HIN_GAI='{0}'
    group by JGYOBU ,NAIGAI ,HIN_GAI
) z on (z.jgyobu = i.jgyobu and z.naigai = i.naigai and z.hin_gai = i.hin_gai)
left outer join (
    select
     HIN_GAI
    ,sum(convert(MI_JITU_QTY,sql_decimal) + convert(SUMI_JITU_QTY,sql_decimal)) qty
    ,count(*) cnt
    from p_sagyo_log
    where jitu_dt >= '{1}'
    and (FROM_SOKO = '92' or TO_SOKO = '96')
//    and RIRK_ID like '4_'
    group by HIN_GAI
) s on (s.hin_gai = i.hin_gai)
left outer join (
    select
     o.HIN_GAI
    ,sum(convert(u.UKEIRE_QTY,sql_decimal)) qty
    ,count(*) cnt
    from p_sukeire u
    inner join p_sshiji_o o 
    on (u.SHIJI_NO = o.SHIJI_NO)
    where u.UKEIRE_DT >= '{1}'
    and o.CANCEL_F <> '1'
    group by
     o.HIN_GAI
) u on (u.hin_gai = i.hin_gai)
where i.HIN_GAI='{0}'
order by i.HIN_NAME desc
""".format(pn, dt8 if dt8 else date.today().strftime("%Y%m%d"))
    item = conn.execute(sql).fetchone()
    return item

def decimal_default_proc(obj):
    if isinstance(obj, Decimal):
        return float(obj)
    print("decimal_default_proc:" + str(obj))
    return obj
    raise TypeError

class DatetimeEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, str):
            return 'a'
        if isinstance(obj, datetime ):
            return obj.strftime("%Y-%m-%d %H:%M:%S")
        if isinstance(obj, date ):
            return obj.strftime("%Y-%m-%d")
        if isinstance(obj, Decimal):
            return float(obj)
        return json.JSONEncoder.default(self, obj)

def print_json(results):
    print('Content-Type:application/json; charset=UTF-8;\n')
    print(json.dumps(results,
#                     default=decimal_default_proc,
                     ensure_ascii=True,
                     indent=4,
                     sort_keys=False,
                     separators=(',', ': '),
                     cls=DatetimeEncoder))

if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        form = cgi.FieldStorage()
        dns = form.getvalue('dns', 'newsdc')
        filename = form.getvalue('filename', '商品化予定.xlsx')
        limit = int(form.getvalue('limit', 0))
        sys.stdout = None
        r = main(dns, filename, limit)
        sys.stdout = sys.__stdout__
        print_json(r)
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("filename", help="商品化予定.xlsx", nargs="?", default="商品化予定.xlsx", type=str)
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--limit", help="default: 0", default=0, type=int)
        args= parser.parse_args()
        r = main(args.dns, args.filename, args.limit)
        print_json(r)
        import pprint
        pprint.pprint(r)
