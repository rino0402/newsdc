# -*- coding: utf-8 -*-
import cgi
import os
import sys
import io
import pyodbc
import json
from decimal import Decimal
from datetime import date, datetime, timedelta

_debug = False
def debug(v):
    if _debug:
        print(v)

def get_req(nm, v):
    form = cgi.FieldStorage()
    if nm in form:
        v = form[nm].value
    return v

def syukadt(dns, order_no, syuka_dt):
    print("syukadt({0}, {1}, {2})".format(dns, order_no, syuka_dt))
    print("pyodbc.connect({0})".format(dns), end=".")
    conn = pyodbc.connect('DSN=' + dns)
    print("ok")
    sql = """
update y_syuka
set SYUKA_YMD='{1}'
,UPD_NOW=left(replace(replace(replace(convert(Now(),SQL_CHAR),'-',''),':',''),' ',''),14)
where JGYOBA='00036003'
and DATA_KBN='1'
and HAN_KBN='2'
and ODER_NO='{0}'
""".format(order_no, syuka_dt)
    print(sql)
    print("conn.execute(update)", end=".")
    conn.execute(sql)
    print("ok")
    sql = "select @@rowcount"
    print("conn.execute({0})".format(sql), end="=")
    rc = conn.execute(sql).fetchone()
    print(rc[0])
    print("conn.commit()", end=".")
    conn.commit()
    print("ok")
    print("conn.close()", end=".")
    conn.close()
    print("ok")

#############
#メイン
#############
if __name__ == "__main__":
#var	url = 'syukadt.py?dns=' + $('#dns').val();
#url += '&ODER_NO=' + $('#eODER_NO').val();
#url += '&SYUKA_YMD=' + $('#eKEY_SYUKA_YMD').val();
    if 'REQUEST_METHOD' in os.environ:
        dns = get_req('dns', 'newsdc')
        syuka_dt = get_req('SYUKA_YMD', '')
        order_no = get_req('ODER_NO', '')
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--ODER_NO", help="ex.20180620", default="", type=str)
        parser.add_argument("--SYUKA_YMD", help="ex.20180620", default="", type=str)
        parser.add_argument("--debug", action="store_true")
        args= parser.parse_args()
        _debug = args.debug
        dns= args.dns
        syuka_dt= args.SYUKA_YMD
        order_no= args.ODER_NO
    syukadt(dns, order_no, syuka_dt)
