# -*- coding: utf-8 -*-
import cgi
import os
import sys
import io
import pyodbc
import json
from decimal import Decimal
from datetime import date, datetime, timedelta
from dateutil import relativedelta

def make(dns, post):
    print("make({0}, {1})".format(dns, post))
    today = datetime.today()
    print(" today={0}".format(today.strftime('%Y-%m-%d')))
    # 今週月曜
    monday = today - relativedelta.relativedelta(days= today.weekday()) 
    print("monday={0}".format(monday.strftime('%Y-%m-%d')))
    # 今週日曜
    sunday = today - relativedelta.relativedelta(days= today.weekday() - 6) 
    print("sunday={0}".format(sunday.strftime('%Y-%m-%d')))
    for i in range(7):
        day = monday + timedelta(days= i)
        print("{0} {1}".format(i, day.strftime('%Y-%m-%d')))

def insupd(dns, tanto_code, caldate, stanto_code, workdet):
    print("insupd1({0}, {1}, {2}, {3}, {4})".format(dns, tanto_code, caldate, stanto_code, workdet))
    print("pyodbc.connect({0})".format(dns), end=".")
    conn = pyodbc.connect('DSN=' + dns)
    print("ok")
    if stanto_code == "":
        sql = """
insert into Sch (
 TANTO_CODE
,Dt
,Holiday
,WorkDet
) values (
 '{0}'
,'{1}'
,''
,'{2}'
)
""".format(tanto_code, caldate, workdet)
        print("insert.", end="")
        result = "insert."
    else:
        sql = """
update Sch set
 WorkDet='{2}'
where TANTO_CODE='{0}'
and Dt='{1}'
""".format(tanto_code, caldate, workdet)
        result = "update."
        
    try:
        conn.execute(sql)
        print("ok")
        result += "ok"
    except pyodbc.ProgrammingError as e:
        print("")
        print("e.type:{0}".format(type(e)))
        for arg in e.args:
            print("e.args:{0}".format(arg))
        result += e.args[1]
    except pyodbc.Error as e:
        if 'SQLExecDirectW' in str(e.args):
            print("W")
        else:
            print("")
            print("e.type:{0}".format(type(e)))
            for arg in e.args:
                print("e.args:{0}".format(arg))
        result += e.args[1]
    except :
        result += "error"

    print("conn.commit()", end=".")
    conn.commit()
    print("ok")

    print("conn.close()", end=".")
    conn.close()
    print("ok")
    return result

def list1(dns, post, stdate):
    print("list1({0}, {1}, {2})".format(dns, post, stdate))
    print("pyodbc.connect({0})".format(dns), end=".")
    conn = pyodbc.connect('DSN=' + dns)
    print("ok")

    today = date.today()
    if stdate != '':
        today = datetime.strptime(stdate, '%Y-%m-%d')
        print(today)
#    print(" today={0}".format(today.strftime('%Y-%m-%d')))
#    monday = today - relativedelta.relativedelta(days= today.weekday()) 
#    print("monday={0}".format(monday.strftime('%Y-%m-%d')))
#    sunday = today - relativedelta.relativedelta(days= today.weekday() - 6) 
#    print("sunday={0}".format(sunday.strftime('%Y-%m-%d')))

    where = "where c.CalDate between '{0}' and '{1}'".format(today.strftime('%Y-%m-%d'), (today + timedelta(6)).strftime('%Y-%m-%d'))
    if post != '':
        where += " and t.POST_CODE = '{0}'".format(post)
    sql = """
select
 t.TANTO_CODE
,t.TANTO_NAME
,SurName(t.TANTO_NAME) SurName
,t.POST_CODE
,t.KUBUN
,t.CalDate
,ifnull(s.TANTO_CODE,'') sTANTO_CODE
,s.Holiday
,s.WorkDet
from (
select
*
from Calendar c , Tanto t
{0}
) t
left outer join Sch s on (t.TANTO_CODE = s.TANTO_CODE and t.CalDate = s.Dt)
order by
 t.TANTO_CODE
,t.CalDate
""".format(where)
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
        TANTO_CODE = form.getvalue('TANTO_CODE', '')
        if TANTO_CODE != '':
            sys.stdout = None
            for i in range(1, 8):
                r["result{0}".format(i)] = insupd(r["dns"],
                       TANTO_CODE,
                       form.getvalue('CalDate{0}'.format(i), ''),
                       form.getvalue('sTANTO_CODE{0}'.format(i), ''),
                       form.getvalue('WorkDet{0}'.format(i), ''))
            sys.stdout = sys.__stdout__
        else:
            r["post"] = form.getvalue('post', '')
            r["stdate"] = form.getvalue('stdate', '')
            sys.stdout = None
            r["data"] = list1(r["dns"], r["post"], r["stdate"])
            sys.stdout = sys.__stdout__
        print('Content-Type:application/json; charset=UTF-8;\n')
#        print(json.dumps(r, default=decimal_default_proc, ensure_ascii=True, indent=4, sort_keys=False, separators=(',', ': ')))
        print(json.dumps(r,  \
                         ensure_ascii=True, indent=4, sort_keys=False, \
                         separators=(',', ': ') \
                         , cls=DatetimeEncoder \
                         ) \
              )
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--post", help="default: ", default="", type=str)
        parser.add_argument("--stdate", help="ex.2018-09-30", default="", type=str)
        parser.add_argument("--debug", action="store_true")
        parser.add_argument("--make", help="空データ作成1週間分",action="store_true")
        args= parser.parse_args()
        if args.make:
            make(args.dns, args.post)
        else:
            import pprint
            pprint.pprint(list1(args.dns, args.post, args.stdate))
