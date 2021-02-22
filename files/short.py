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
import openpyxl
import traceback
from datetime import date, datetime, timedelta
from decimal import Decimal
#----------------------------------------------------------------
def main(r):
    print("main({})".format(r))
    if r["jgyobu"] == '7':
        r["jcode"] = '00023210'
    elif r["jgyobu"] == 'D':
        r["jcode"] = '00023510'
    elif r["jgyobu"] == '4':
        r["jcode"] = '00023410'
    elif r["jgyobu"] == '5':
        r["jcode"] = '00021397'
    print("pyodbc.connect({0})".format(r["dns"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dns"])
    print("ok")
    r["list"] = short_list(conn, r)
    print("conn.close()", end=".")
    conn.close()
    print("ok")
    return r
#----------------------------------------------------------------
def short_list(conn, r):
    print("short_list({})".format(r))
    sql = "select {}".format(' top {}'.format(r["limit"]) if r["limit"] > 0 else '')
    sql += """
 i.HIN_GAI
,i.HIN_NAI
,i.HIN_NAME
,GetSupplyNm(i.NAI_BUHIN) Nai
,GetSupplyNm(i.GAI_BUHIN) Gai
,z.qty
,round(ifnull(convert(a.AVE_SYUKA,sql_decimal),0),1) AveSyuka
,if(ifnull(convert(a.AVE_SYUKA,sql_decimal),0) = 0
   ,null()
   ,round(ifnull(z.qty,0)/convert(a.AVE_SYUKA,sql_decimal),1)
 ) ZMonth
,if(if(ifnull(convert(a.AVE_SYUKA,sql_decimal),0) = 0
      ,null()
      ,round(ifnull(z.qty,0)/convert(a.AVE_SYUKA,sql_decimal),1)
    )<=1
   ,'○'
   ,''
   )
Month1
,if(if(ifnull(convert(a.AVE_SYUKA,sql_decimal),0) = 0
	  ,null()
	  ,round(ifnull(z.qty,0)/convert(a.AVE_SYUKA,sql_decimal),1)
	  )<=(5/30)
   ,'○'
   ,''
   )
Day5
,pn.NaiDisconYm
,pns.Biko
,pnx.Biko PnxBiko
,g.Qty
,g.YoteiDt
from item i
left outer join (
	select
	 HIN_GAI
	,sum(convert(YUKO_Z_QTY,sql_decimal))	qty
	from Zaiko
	where JGYOBU='{0}'
	  and NAIGAI='1'
	group by HIN_GAI
) z
	on (z.HIN_GAI=i.HIN_GAI)
left outer join ave_syuka a
	on (a.JGYOBU='{0}' and a.NAIGAI='1' and i.HIN_GAI=a.HIN_GAI)
inner join PnNew pn
	on (pn.JCode='00036003' and pn.ShisanJCode='{1}' and i.HIN_GAI=pn.Pn)
left outer join PnShort pns
	on (pns.JCode='{1}' and i.HIN_GAI=pns.Pn)
left outer join ShortXls pnx
	on (i.HIN_GAI=pnx.Pn)
left outer join (
select 
 Pn
,Qty
,YoteiDt
from GOrder
where Pn <> ''
and Qty <> 0
and YoteiDt <> ''
union
select
 HIN_GAI Pn
,Convert(N_YOTEI_QTY,sql_decimal) Qty
,SUBSTRING(N_YOTEI_DT,1,4)
+'-'
+SUBSTRING(N_YOTEI_DT,5,2)
+'-'
+SUBSTRING(N_YOTEI_DT,7,2)
YoteiDt
from PLN_Y_NYUKA
where Pn <> ''
and Qty <> 0
and N_YOTEI_DT > '20200000'
) g
    on (z.HIN_GAI = g.Pn)
where i.JGYOBU='{0}' and i.NAIGAI='1'
and (z.qty > 0 or ifnull(convert(a.AVE_SYUKA,sql_decimal),0) > 0)
order by
 Day5 desc
,Month1 desc
,if(ifnull(convert(a.AVE_SYUKA,sql_decimal),0) = 0,1,0)
,ZMonth
,AveSyuka desc
,z.qty desc
,i.HIN_GAI
,g.YoteiDt
""".format(r["jgyobu"], r["jcode"])
    rlist = []
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
            if len(rlist) > 0 and rlist[-1]["HIN_GAI"] == d["HIN_GAI"]:
                print(d["HIN_GAI"])
                n += 1
                rlist[-1]["Qty{}".format(n)] = d["Qty"]
                rlist[-1]["YoteiDt{}".format(n)] = d["YoteiDt"]
            else:
                n = 1
                rlist.append(d)
    except pyodbc.ProgrammingError as e:
        print(e)
        raise
    except pyodbc.Error as e:
        print(e)
        raise
    except:
        print('error')
        raise
    return rlist
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
def excel(r):
    print("excel()")
    results = {}
    xls = os.path.dirname(os.path.abspath(__file__))
    xls = os.path.abspath(xls + "\\short.xlsx")
    print(xls)
    wb = openpyxl.load_workbook(xls)
    print(str(wb.sheetnames))
    sheet = wb['在庫検討data']
    row = 2
    for d in r["list"]:
        sheet["A{}".format(row)] = d["HIN_GAI"]
        sheet["B{}".format(row)] = d["HIN_NAI"]
        sheet["C{}".format(row)] = d["HIN_NAME"]
        sheet["D{}".format(row)] = d["Nai"]
        sheet["E{}".format(row)] = d["Gai"]
        sheet["F{}".format(row)] = d["qty"]
        sheet["G{}".format(row)] = d["AveSyuka"]
        sheet["H{}".format(row)] = d["ZMonth"]
        sheet["I{}".format(row)] = d["Month1"]
        sheet["J{}".format(row)] = d["Day5"]
        sheet["K{}".format(row)] = d["NaiDisconYm"]
        sheet["L{}".format(row)] = d["Biko"]
        sheet["M{}".format(row)] = d["PnxBiko"]
        try:
            sheet["N{}".format(row)] = "{0}\n{1}".format(d["YoteiDt"][-5:].replace("-","/"), int(d["Qty"]))
        except:
            sheet["N{}".format(row)] = ""
        try:
            sheet["O{}".format(row)] = "{0}\n{1}".format(d["YoteiDt2"][-5:].replace("-","/"), int(d["Qty2"]))
        except:
            sheet["O{}".format(row)] = ""
        row += 1
    sheet.delete_rows(row, 65536)
    #保存
    results["excel"] = "short_{}.xlsx".format(datetime.now().strftime("%Y%m%d%H%M%S_%f"))
    print(results["excel"])
    xls_sv = os.path.dirname(os.path.abspath(__file__))
    xls_sv = os.path.abspath(xls_sv + "\\" + results["excel"])
    wb.save(xls_sv)
    return results
#----------------------------------------------------------------
if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        form = cgi.FieldStorage()
        r = {}
        r["dns"] = form.getvalue('dns', 'newsdc')
        r["limit"] = int(form.getvalue('limit', 0))
        r["jgyobu"] = form.getvalue('jgyobu', '7')
        r["jcode"] = form.getvalue('jcode', '00023210')
        sys.stdout = None
        r = main(r)
        if form.getvalue('excel', '') == "1":
            r = excel(r)
        sys.stdout = sys.__stdout__
        print_json(r)
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--jgyobu", help="default: ", default="7", type=str)
        parser.add_argument("--jcode", help="default: ", default="00023210", type=str)
        parser.add_argument("--limit", help="default: 0", default=0, type=int)
        parser.add_argument("--excel", help="make short.xls", action="store_true")
        r = main(vars(parser.parse_args()))
        if r["excel"]:
            r = excel(r)
        import pprint
        pprint.pprint(r)
