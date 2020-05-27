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

def main(dns, pallet_no1, pallet_no2, id_no, case_qty, limit):
    print("main({0}, {1}, {2}, {3}, {4}, {5})".format(dns, pallet_no1, pallet_no2, id_no, case_qty, limit))
    results = {}
    results["dns"] = dns
    results["limit"] = limit
    results["id_no"] = id_no
    results["case_qty"] = case_qty
    results["pallet_no1"] = pallet_no1
    results["pallet_no2"] = pallet_no2

    print("pyodbc.connect({0})".format(dns), end=".")
    conn = pyodbc.connect('DSN=' + dns)
    print("ok")

    if id_no != "":
        results["update"] = y_syuka_update(conn, pallet_no1, case_qty, id_no)
        conn.commit()
    else:
        results["list"] = y_syuka_list(conn, pallet_no1, pallet_no2, limit)

    print("conn.close()", end=".")
    conn.close()
    print("ok")

    return results

def y_syuka_update(conn, pallet_no, case_qty, id_no):
    print("y_syuka_update({0}, {1}, {2})".format(pallet_no, case_qty, id_no))
    sql = """
update y_syuka
set LK_SEQ_NO='{0}'
,KENPIN_SURYO='{1}'
,UPD_NOW='{3}'
where KEY_ID_NO='{2}'
""".format(pallet_no, case_qty, id_no, datetime.now().strftime("%Y%m%d%H%M%S"))
    print(sql)
    try:
        conn.execute(sql)
        return conn.execute("select @@rowcount").fetchone()[0]
    except pyodbc.ProgrammingError as e:
        print(e)
        raise
    except pyodbc.Error as e:
        print(e)
        raise
    except:
        print('error')
        raise

"""
荷積明細
"""
def y_syuka_list(conn, pallet_no1, pallet_no2, limit):
    print("y_syuka_list({0}, {1}, {2})".format(pallet_no1, pallet_no2, limit))
    sql = """
select {}
*
from y_syuka y
left outer join acorder a on (y.ODER_NO = a.IdNo)
""".format(' top {}'.format(limit) if limit > 0 else '')

    sql += "where y.key_muke_code='00036003'"
    sql += " and y.key_syuka_ymd > '20200101'"
    if pallet_no2 != "":
        sql += " and y.LK_SEQ_NO between '{}' and '{}'".format(pallet_no1, pallet_no2)
    elif '%' in pallet_no1:
        sql += " and y.LK_SEQ_NO like '{}'".format(pallet_no1)
    elif pallet_no1 != "":
        sql += " and y.LK_SEQ_NO = '{}'".format(pallet_no1)
    sql += """    
order by
 if(y.LK_SEQ_NO = '','99999999',y.LK_SEQ_NO)
,if(y.LK_SEQ_NO = '',Null(),convert(y.KENPIN_SURYO,sql_decimal))
,if(y.LK_SEQ_NO = '',Null(),convert(y.SURYO,sql_decimal))
,if(y.KENPIN_YMD in ('','00000000'),'99999999' + convert(ifnull(a.Row,'999999'),sql_char), y.KENPIN_YMD + y.KENPIN_HMS)
"""
    
    r = []
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
            d = dict(zip(columns, row))
            if len(r) > 0 and r[-1]["KEY_ID_NO"] == d["KEY_ID_NO"]:
                pass
            else:
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

def decimal_default_proc(obj):
    if isinstance(obj, Decimal):
        return float(obj)
    raise TypeError

def print_json(results):
    print('Content-Type:application/json; charset=UTF-8;\n')
    print(json.dumps(results,
                     default=decimal_default_proc,
                     ensure_ascii=True,
                     indent=4,
                     sort_keys=False,
                     separators=(',', ': ')))

if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        form = cgi.FieldStorage()
        dns = form.getvalue('dns', 'newsdc')
        pallet_no1 = form.getvalue('pallet_no1', '')
        pallet_no2 = form.getvalue('pallet_no2', '')
        id_no = form.getvalue('id_no', '')
        case_qty = form.getvalue('case_qty', '')
        limit = int(form.getvalue('limit', 0))
        sys.stdout = None
        r = main(dns, pallet_no1, pallet_no2, id_no ,case_qty , limit)
        sys.stdout = sys.__stdout__
        print_json(r)
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--id_no", help="ID_No", default="", type=str)
        parser.add_argument("--case_qty", help="Case_Qty", default="", type=str)
        parser.add_argument("--limit", help="default: 0", default=0, type=int)
        parser.add_argument("pallet_no", help="P/T No.", default="", nargs='*')
        args= parser.parse_args()
        print(args)
        pallet_no1 = args.pallet_no[0]
        try:
            pallet_no2 = args.pallet_no[1]
        except:
            pallet_no2 = ""
        r = main(args.dns, pallet_no1, pallet_no2, args.id_no, args.case_qty, args.limit)
        import pprint
        pprint.pprint(r)
