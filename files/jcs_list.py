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

def main(dns, limit):
    print("main({0}, {1})".format(dns, limit))
    results = {}
    results["dns"] = dns
    results["limit"] = limit
    print("pyodbc.connect({0})".format(dns), end=".", flush=True)
    conn = pyodbc.connect('DSN=' + dns)
    print(conn.getinfo(pyodbc.SQL_DATA_SOURCE_NAME), flush=True)

    results["data"] = load(conn, limit)

    print("conn.close()", end=".", flush=True)
    conn.close()
    print("ok", flush=True)

    return results

def load(conn, limit):
    print("load({0}, {1})".format(conn.getinfo(pyodbc.SQL_DATA_SOURCE_NAME), limit))
    sql = """
select
 i.NohinNo
,i.NohinNo2
,i.DestCode
,i.DlvDt
,i.Location
,i.MazdaPn
,i.Qty
,n.Pn
,n.EntID
,n.EntTm
,p.HAKKO_DT
,p.HIN_CHECK_TANTO
,p.HIN_CHECK_DATETIME
,p.HIN_CHECK_LABEL_CNT
,p.HIN_CHECK_GENPIN_CNT
,p.KAN_F
,p.KAN_DT
,p.BUNNOU_CNT
,p.UKEIRE_QTY
,z.currQty
from JcsIdo i
left outer join JcsNohin n
 on (i.NohinNo = n.NohinNo and i.NohinNo2 = n.NohinNo2
 and i.DestCode = n.DestCode and i.DlvDt = n.DlvDt)
left outer join P_SSHIJI_O p
 on ((rtrim(i.NohinNo) + i.NohinNo2) = p.SHIJI_NO)
left outer join JcsZaiko z
 on (n.Pn = z.Pn)
where i.NohinNo <> ''
and i.DlvDt >= CURDATE()
order by
 i.DlvDt
,i.DestCode
,i.Location
,i.NohinNo
"""
    ### Pos指図票
    sql = """
select
 p.SHIJI_NO
,left(p.SHIJI_NO,6)		NohinNo
,SUBSTRING(p.SHIJI_NO,7,1)	NohinNo2
,p.BIKOU			DestCode
,p.ORDER_DT			DlvDt
,p.F_CLASS_CODE			Location
,p.HIN_GAI			MazdaPn
,p.SHIJI_QTY 			Qty
,p.S_CLASS_CODE			Pn
,n.EntID
,n.EntTm
,p.HAKKO_DT
,p.HIN_CHECK_TANTO
,p.HIN_CHECK_DATETIME
,p.HIN_CHECK_LABEL_CNT
,p.HIN_CHECK_GENPIN_CNT
,p.KAN_F
,p.KAN_DT
,p.BUNNOU_CNT
,p.UKEIRE_QTY
,z.currQty
,p.CANCEL_F
,p.CANCEL_DATETIME
,p.TANTO_CODE
,p.SHONIN_CODE
,p.PRI_SHIJI
,p.N_CLASS_CODE
,p.PRINT_DATETIME
,i.Dt
,i.DestCode iDestCode
,i.DestName
,i.IQty
,itm.NameE
,itm.NameJ
,itm.SSpec
,itm.SType
,itm.GPn
,pos_i.K_KEITAI
,pos_i.PACKING_NO
,pos_i.CATEGORY_CODE
from P_SSHIJI_O p
left outer join JcsNohin n
 on ((rtrim(n.NohinNo) + n.NohinNo2) = p.SHIJI_NO)
left outer join JcsZaiko z
 on (p.S_CLASS_CODE = z.Pn)
left outer join JcsIdo i
 on (p.S_CLASS_CODE = i.Pn and convert(i.IQty,sql_decimal) > 0)
left outer join JcsItem itm
 on (p.HIN_GAI = itm.MazdaPn)
left outer join Item pos_i
 on (p.JGYOBU = pos_i.JGYOBU and p.NAIGAI = pos_i.NAIGAI and p.HIN_GAI = pos_i.HIN_GAI)
where p.ORDER_DT > {}
order by
 DlvDt
,DestCode
,Location
,p.SHIJI_NO
,i.Id desc
""".format(date.today().strftime('%Y%m%d'))
    r = []
    try:
        cursor = conn.cursor()
        cursor.execute(sql)
        columns = [column[0] for column in cursor.description]
        prev = ""
        for row in cursor.fetchall():
            for i, col in enumerate(row):
                if isinstance(col, str):
                    row[i] = col.rstrip()
#                if col is None:
#                    row[i] = ""
            if row.SHIJI_NO != prev:
                r.append(dict(zip(columns, row)))
            prev = row.SHIJI_NO
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
        limit = int(form.getvalue('limit', 0))
        sys.stdout = None
        r = main(dns, limit)
        sys.stdout = sys.__stdout__
        print_json(r)
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--limit", help="default: 0", default=0, type=int)
        args= parser.parse_args()
        r = main(args.dns, args.limit)
        print_json(r)
        import pprint
        pprint.pprint(r)
