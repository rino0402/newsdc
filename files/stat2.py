# -*- coding: utf-8 -*-
import cgi
import os
import sys
import io
import pyodbc
import json
from decimal import Decimal
from datetime import date, datetime, timedelta

def stat2(dns):
    print("stat2({0})".format(dns))
    print("pyodbc.connect({0})".format(dns), end=".")
    conn = pyodbc.connect('DSN=' + dns)
    print("ok")
    sql = """
select
 h.JCode
,if(h.Soko = 'ACD','ZACS',h.Soko) Soko
,h.KJCode
,h.SJCode
,rtrim(j.Name) + if(h.Soko = 'ACD','(湖南)','') Name
,count(distinct h.Pn) cntPn
,count(*) cnt
,sum(h.Qty) sumQty
,count(distinct if(h.CyuKbn in ('2','3'),h.Pn,null())) cntPn1
,sum(if(h.CyuKbn in ('2','3'),1,0)) cnt1
,sum(if(h.CyuKbn in ('2','3'),h.Qty,0)) sumQty1
,count(distinct if(h.CyuKbn in ('2','3'),null(),h.Pn)) cntPn2
,sum(if(h.CyuKbn in ('2','3'),0,1)) cnt2
,sum(if(h.CyuKbn in ('2','3'),0,h.Qty)) sumQty2
from HMTAH015 h
left outer join JGyobu j
on (j.JCode = h.SJCode)
where h.Stts='2' //and h.Soko<>'ACD'
group by
 h.JCode
,h.Soko
,h.KJCode
,h.SJCode
,j.Name
union
select
 'Z0036003'
,'ZACD' Soko
,'00025800'
,'00025800'
,'エアコン(緊急)'
,count(distinct Pn) cntPn
,count(*) cnt
,sum(Qty) sumQty
,count(distinct Pn) cntPn1
,count(*) cnt1
,sum(Qty) sumQty1
,Null cntPn2
,Null cnt2
,Null sumQty2
from AcShort
group by
 '00036003'
,Soko
,'00025800'
,'エアコン(緊急)'
order by
 1
,2
,3
,4
"""
    print(sql)
    print("conn.execute(sql)", end=".")
    cursor = conn.cursor()
    cursor.execute(sql)
    columns = [column[0] for column in cursor.description]
    data = []
    for row in cursor.fetchall():
        for i, v in enumerate(row):
            if v is None:
                row[i]=""
            elif isinstance(v, str):
                row[i]=v.rstrip()
        dic = dict(zip(columns, row))
        data.append(dic)
    print("ok")
    print("conn.close()", end=".")
    conn.close()
    print("ok")
    return data

def decimal_default_proc(obj):
    if isinstance(obj, Decimal):
        return float(obj)
    raise TypeError

#############
#メイン
#############
if __name__ == "__main__":
#var	url = 'syukadt.py?dns=' + $('#dns').val();
#url += '&ODER_NO=' + $('#eODER_NO').val();
#url += '&SYUKA_YMD=' + $('#eKEY_SYUKA_YMD').val();
    if 'REQUEST_METHOD' in os.environ:
        form = cgi.FieldStorage()
        r = {}
        r["dns"] = form.getvalue('dns', 'newsdc')
        sys.stdout = None
        r["data"] = stat2(r["dns"])
        sys.stdout = sys.__stdout__
        print('Content-Type:application/json; charset=UTF-8;\n')
        print(json.dumps(r, default=decimal_default_proc, ensure_ascii=True, indent=4, sort_keys=False, separators=(',', ': ')))
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--debug", action="store_true")
        args= parser.parse_args()
        dns= args.dns
        import pprint
        pprint.pprint(stat2(args.dns))
