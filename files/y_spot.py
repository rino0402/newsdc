# -*- coding: utf-8 -*-

import os
import sys
import io
import cgi
import re
import datetime
import xlrd
import pyodbc
import json
from decimal import Decimal
#from datetime import datetime, timedelta

def main(dns, soko, limit):
    print("main({0}, {1}, {2})".format(dns, soko, limit))
    results = {}
    print("pyodbc.connect({0})".format(dns), end=".", flush=True)
    conn = pyodbc.connect('DSN=' + dns)
    print("ok", flush=True)

    sql = """
select
 h.SyukaDt1 SyukaDt
,h.Stts1 Stts
,case h.Aitesaki1
 when '22000443' then '東日本サテ'
 when '22000444' then '西日本サテ'
 when '22000446' then '福岡サテ'
 else h.AitesakiName1
 end aName
,case h.Aitesaki2
 when '22000443' then '東日本サテ'
 when '22000444' then '西日本サテ'
 when '22000446' then '福岡サテ'
 else h.AitesakiName2
 end aName2
,i.ST_SOKO
,sk.SOKO_NAME
,h.JGYOBU
,h.Pn
,h.Qty
,h.Cnt
,z.*
,i.HIN_NAME
,if(i.ST_SOKO <> left(sk.SOKO_NAME,2),i.ST_SOKO + ' ','') + ifnull(sk.SOKO_NAME,i.ST_SOKO) SokoName
from (
select
 j.JGYOBU
,h.Pn
,sum(h.Qty) Qty
,count(*) Cnt
,min(h.Stts) Stts1
,max(h.Stts) Stts2
,min(h.SyukaDt) SyukaDt1
,max(h.SyukaDt) SyukaDt2
,min(h.Aitesaki) Aitesaki1
,max(h.Aitesaki) Aitesaki2
,min(if(h.Aitesaki like '2200044_',null,h.AitesakiName)) AitesakiName1
,max(if(h.Aitesaki like '2200044_',null,h.AitesakiName)) AitesakiName2
from HMTAH015 h
inner join JGyobu j on (h.SJCode = j.JCode)
where h.Soko in (select Value from Config where Name = 'Soko')
and h.Aitesaki not in ('00039171')
and h.Stts in ('3','4')
and h.CyuKbn in ('2','3')
and h.IDNo not in (select distinct KEY_ID_NO from y_syuka where KAN_KBN = '9')
group by
 j.JGYOBU
,h.Pn
) h
left outer join (
select 
 JGYOBU
,HIN_GAI
,Sum(convert(YUKO_Z_QTY,sql_decimal)) QtyZ
,Sum(if(Soko_No<>'92' and GOODS_ON= '0',convert(YUKO_Z_QTY,sql_decimal),0)) Qty0
,Sum(if(Soko_No<>'92' and GOODS_ON<>'0',convert(YUKO_Z_QTY,sql_decimal),0)) Qty1
,Sum(if(Soko_No='92',convert(YUKO_Z_QTY,sql_decimal),0)) Qty92
from zaiko
group by
 JGYOBU
,HIN_GAI
) z on (h.JGYOBU = z.JGYOBU and h.Pn = z.HIN_GAI)
left outer join item i on (h.JGYOBU = i.JGYOBU and i.NAIGAI = '1' and h.Pn = i.HIN_GAI)
left outer join Soko sk on (i.ST_SOKO = sk.Soko_No)
where h.Qty >= z.Qty0
order by
 SyukaDt
,Stts desc
,i.ST_SOKO
"""
    try:
        cursor = conn.cursor()
        cursor.execute(sql)
        columns = [column[0] for column in cursor.description]
#        print(columns)
        data = []
        for row in cursor.fetchall():
            for i, col in enumerate(row):
                if isinstance(col, str):
                    row[i] = col.rstrip()
            data.append(dict(zip(columns, row)))
#        print(data)
        results["data"] = data
    except pyodbc.ProgrammingError as e:
        print(e)
        raise
    except pyodbc.Error as e:
        print(e)
        raise
    except:
        raise

    print("conn.close()", end=".", flush=True)
    conn.close()
    print("ok", flush=True)
    return results

def print_json(results):
    print('Content-Type:application/json; charset=UTF-8;\n')
    print(json.dumps(results,
                     default=decimal_default_proc,
                     ensure_ascii=True,
                     indent=4,
                     sort_keys=False,
                     separators=(',', ': ')))

def decimal_default_proc(obj):
    if isinstance(obj, Decimal):
        return float(obj)
    raise TypeError

if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        form = cgi.FieldStorage()
        dns = form.getvalue('dns', 'newsdc')
        soko = form.getvalue('Soko', '')
        limit = form.getvalue('limit', 0)
        sys.stdout = None
        r = main(dns, soko, limit)
        sys.stdout = sys.__stdout__
        print_json(r)
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--soko", help="CSD ECD", default="", type=str)
        parser.add_argument("--limit", help="default: 0", default=0, type=int)
        args= parser.parse_args()
        print("{0} 出荷商品化待ち {1}".format(__file__, args))
        r = main(args.dns, args.soko, args.limit)
        from pprint import pprint
        pprint(r)
