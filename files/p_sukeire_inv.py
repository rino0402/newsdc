# -*- coding: utf-8 -*-

import cgi
import os
import sys
import io
import datetime
import pyodbc
import json
import openpyxl
from decimal import Decimal

# ----------------------------------------------------------------
def decimal_default_proc(obj):
    if isinstance(obj, Decimal):
        return float(obj)
    raise TypeError
# ----------------------------------------------------------------
def get_req(nm, v):
    if nm in form:
        v = form[nm].value
    return v
# ----------------------------------------------------------------
def get_dns():
    dns = 'newsdc'
    if 'HTTP_HOST' not in os.environ:
        dns = 'newsdc4'
    elif os.environ['HTTP_HOST'] == 'w0':
        dns = 'newsdc4'
    return get_req('dns',dns)
# ----------------------------------------------------------------
def get_where(w,nm,v1,v2):
    if(v1 == ''):
        return w
    if(w == ''):
        w = ' where '
    else:
        w += ' and '
    if(nm == 'u.UKEIRE_DT' or nm == 'UKEIRE_DT'):
        w += nm + " between '" + v1 + "' and '" + v2 + "'"
    else:
        w += nm + " = '" + v1 + "'"
    return w
# ----------------------------------------------------------------
# レスポンス
# ----------------------------------------------------------------
def print_response(r):
    print('Content-Type:application/json; charset=UTF-8;\n\n')
    print(json.dumps(r, default=decimal_default_proc, ensure_ascii=True, indent=4, sort_keys=False, separators=(',', ': ')))
    log('▽')
    return
# ----------------------------------------------------------------
# get_browser
# ----------------------------------------------------------------
def get_browser():
    agent = os.environ['HTTP_USER_AGENT'] if 'HTTP_USER_AGENT' in os.environ else ''
    agent = agent.lower()
    if 'msie' in agent:
        browser = 'IEolder'
    elif 'trident' in agent:
        browser = 'IE11'
    elif 'edge' in agent:
        browser = 'Edge'
    elif 'chrome' in agent:
        browser = 'Chrome'
    elif 'firefox' in agent:
        browser = 'FireFox'
    elif 'opera' in agent:
        browser = 'Opera'
    else:
        browser = 'unknown'
    return browser
# ----------------------------------------------------------------
# log
# ----------------------------------------------------------------
def log(t):
    now = now = datetime.datetime.now()
    s = now.strftime("%Y/%m/%d %H:%M:%S")
    s += ':' + os.environ['REMOTE_ADDR'] if 'REMOTE_ADDR' in os.environ else ''
    s += ':' + get_browser()
    s += ':' + t
    s += ':' + os.path.basename(__file__)
    s += '' + os.environ['QUERY_STRING'] if 'QUERY_STRING' in os.environ else ''
    filename = os.path.dirname(os.path.abspath(__file__)) + '\\p_sukeire_inv.log'
    f = open(filename, 'a')
    f.write(s + '\n')
    f.close()
# ----------------------------------------------------------------
def main(dns, dt1, dt2, shimuke, tanto, check):
    log('△')
    results = {}
    results["HTTP_HOST"] = os.environ['HTTP_HOST'] if 'HTTP_HOST' in os.environ else ''
    results["REQUEST_URI"] = os.environ['REQUEST_URI'] if 'REQUEST_URI' in os.environ else ''
    results["dns"] = dns
    today = datetime.date.today()
    today_def = today.strftime("%Y%m%d")
    results["UKEIRE_DT1"] = dt1 # get_req('UKEIRE_DT1',today_def)
    results["UKEIRE_DT2"] = dt2 # get_req('UKEIRE_DT2',today_def)
    results["SHIMUKE_CODE"] = shimuke # get_req('SHIMUKE_CODE','')
    results["S_TANTO"] = tanto # get_req('S_TANTO','')
    results["CHECK"] = check # get_req('CHECK','')
    try:
        c = 'DSN=' + dns # get_dns()
        conn = pyodbc.connect(c)
    except:
        results["error"] = 'error:pyodbc.connect():' + c
        print_response(results)
        sys.exit()
    cursor = conn.cursor()
    where = "where u.UKEIRE_QTY <> 0"
    where = get_where(where,'u.UKEIRE_DT', dt1, dt2)
    where = get_where(where,'u.SHIMUKE_CODE', shimuke, '')
    where = get_where(where,'o.S_TANTO', tanto, '')
    sql = """
    select
    //top 1
     u.SHIJI_NO
    ,u.SEQNO
    ,o.CANCEL_F
    ,u.SHIMUKE_CODE
    ,rtrim(o.S_TANTO)   S_TANTO
    ,u.UKEIRE_DT
    ,o.JGYOBU
    ,o.NAIGAI
    ,rtrim(o.HIN_GAI)   HIN_GAI
    ,rtrim(i.HIN_NAME)  HIN_NAME
    ,rtrim(i.L_KISHU1)  L_KISHU1
    ,CEILING(convert(o.SHIJI_QTY,sql_decimal))   SHIJI_QTY
    ,CEILING(convert(u.UKEIRE_QTY,sql_decimal))  UKEIRE_QTY
    ,convert(i.S_KOUSU_BAIKA,sql_decimal)       KoryoPrc
    ,convert(u.UKEIRE_QTY,sql_decimal)
    *convert(i.S_KOUSU_BAIKA,sql_decimal)       KoryoAmt
    ,convert(i.S_SHIZAI_BAIKA,sql_decimal)      HakoPrc
    ,convert(u.UKEIRE_QTY,sql_decimal)
    *convert(i.S_SHIZAI_BAIKA,sql_decimal)      HakoAmt
    ,convert(i.S_GAISO_TANKA,sql_decimal)       GaisoPrc
    ,convert(u.UKEIRE_QTY,sql_decimal)
    *convert(i.S_GAISO_TANKA,sql_decimal)       GaisoAmt
    ,convert(i.S_PPSC_KAKO_KOSU,sql_decimal)    KakoPrc
    ,convert(u.UKEIRE_QTY,sql_decimal)
    *convert(i.S_PPSC_KAKO_KOSU,sql_decimal)    KakoAmt
    ,convert(i.S_BU_KAKO_KOSU,sql_decimal)      BuPrc
    ,convert(u.UKEIRE_QTY,sql_decimal)
    *convert(i.S_BU_KAKO_KOSU,sql_decimal)      BuAmt
    """
    if (check == '1' ): # get_req('CHECK','')
        sql += """
    ,round(
     CEILING((
     convert(i.BEF_KOUTEI_10,sql_decimal)
    +round((                                        //作業時間:計算
    +(convert(i.SEI_LABEL_QTY,sql_decimal) * 4)     //ラベル貼り
    +c.KosoTm                                       //個装作業
    +(c.DokonTm * 4)                                //同梱作業
    +c.SKakoTm                                      //加工作業
    +c.SyugoTm                                      //集合梱包
    )*1.15,0)
    +convert(i.AFT_KOUTEI_10,sql_decimal)
    +convert(i.PLUS_KOUSU,sql_decimal)
    )/6)/10
    *convert(i.SEI_RATE,sql_decimal)                //分レート
    ,2)
    cKoryoPrc                                       //工料
    ,ifNull(c.HakoPrc,0)	cHakoPrc
    ,ifNull(c.GaisoPrc,0)	cGaisoPrc
    ,ifNull(c.KakoPrc,0)	cKakoPrc
    ,ifNull(c.BuPrc,0)	cBuPrc
    ,ifNull(c.HikitoriTm,0)	cHikitoriTm
    """
    sql += """
    From P_SUKEIRE u
    left outer join P_SSHIJI_O o on (u.shiji_no=o.shiji_no)
    left outer join Item i on (o.JGYOBU = i.JGYOBU and o.NAIGAI = i.NAIGAI and o.HIN_GAI = i.HIN_GAI)
    """
    if (check == '1' ): # get_req('CHECK','') 
        sql += """
    left outer join (
    select
     k.SHIMUKE_CODE
    ,k.JGYOBU
    ,k.NAIGAI
    ,k.HIN_GAI
    ,sum(if(k.DATA_KBN='1' or k.KO_SYUBETSU in ('03','90'),convert(s.S_KOUSU,sql_decimal) * convert(k.KO_QTY,SQL_DECIMAL),0))
    KosoTm
    ,sum(if(k.DATA_KBN='1' or k.KO_SYUBETSU in ('03','90'),convert(s.SEI_SYU_KON,sql_decimal),0))
    SyugoTm
    ,sum(if(k.DATA_KBN='3' and k.KO_SYUBETSU in ('01'),convert(k.KO_QTY,SQL_DECIMAL),0))
    DokonTm
    ,sum(if(k.DATA_KBN='3' and k.KO_SYUBETSU in ('05'),convert(s.S_KOUSU,sql_decimal) * convert(k.KO_QTY,SQL_DECIMAL),0))
    SKakoTm
    ,sum(if((k.DATA_KBN='1' or k.KO_SYUBETSU in ('03','90')) and s.SEI_KBN not in ('1','2'),convert(s.G_ST_URITAN,sql_decimal) * convert(k.KO_QTY,SQL_DECIMAL),0))
    HakoPrc
    ,sum(if((k.DATA_KBN='2' or k.KO_SYUBETSU in ('91')) and s.SEI_KBN not in ('1','2')
        ,Gaiso(convert(k.KO_QTY,SQL_DECIMAL),convert(s.G_ST_URITAN,sql_decimal))
        ,0))
    GaisoPrc
    ,sum(if(k.DATA_KBN='3' and k.KO_SYUBETSU in ('06','04')
        ,Kako(convert(k.KO_QTY,SQL_DECIMAL),convert(s.S_KOUSU,sql_decimal))
        ,0))
    KakoPrc
    ,sum(if(k.DATA_KBN='3' and k.KO_SYUBETSU in ('07')
        ,Kako(convert(k.KO_QTY,SQL_DECIMAL),convert(s.S_KOUSU,sql_decimal))
        ,0)
        )
    BuPrc
    //引取
    ,sum(if(k.DATA_KBN='3' and k.KO_SYUBETSU in ('07') and k.KO_BIKOU like '引%取%' and ISNUMERIC(k.KO_HIN_GAI) = 1
            ,convert(k.KO_HIN_GAI,SQL_DECIMAL)
            ,0)
        )
    HikitoriTm
    From P_COMPO_K k
    left outer join ITEM s
     on (k.KO_JGYOBU =s.JGYOBU and k.KO_NAIGAI =s.NAIGAI and k.KO_HIN_GAI =s.HIN_GAI)
    where k.DATA_KBN in ('1','2','3')
    //検索速度をあげる為
     and k.HIN_GAI in (
     select distinct HIN_GAI from P_SSHIJI_O
     where shiji_no in (
     select distinct shiji_no from P_SUKEIRE
    """
        check_w = ""
        check_w = get_where(check_w,'UKEIRE_DT', dt1, dt2)
        check_w = get_where(check_w,'SHIMUKE_CODE', shimuke, '')
    #    check_w = get_where(check_w,'S_TANTO',get_req('S_TANTO',''),'')
        sql += check_w
        sql += """
     ))
    group by
     k.SHIMUKE_CODE
    ,k.JGYOBU
    ,k.NAIGAI
    ,k.HIN_GAI
    ) c on (o.SHIMUKE_CODE = c.SHIMUKE_CODE and o.JGYOBU = c.JGYOBU and o.NAIGAI = c.NAIGAI and o.HIN_GAI = c.HIN_GAI)
    """
    sql += where
    sql += " order by u.UKEIRE_DT,o.HIN_GAI,u.SHIJI_NO,u.SEQNO"
    #print(sql)
    results["sql"] = sql
    try:
        cursor.execute(sql)
    except:
        results["error"] = 'error:cursor.execute()'
        return results
        # print_response(results)
        # sys.exit()
    data = []
    columns = [column[0] for column in cursor.description]
    for c in cursor.fetchall():
        data.append(dict(zip(columns, c)))
    cursor.close()
    conn.close()
    results["data"] = data
    return results
# ----------------------------------------------------------------
def rfbill(r):
    print("rfbill(r:{})".format(len(r)))
    xls = os.path.dirname(os.path.abspath(__file__))
    xls = os.path.abspath(xls + "\\rfbill.xlsx")
    print(xls)

    wb = openpyxl.load_workbook(xls)
    print(str(type(wb)))
    print(str(wb.sheetnames))
    sheet = wb['RFSP']
    sheet['I1'] = '2020/3/18'
    sheet['A8'] = '2020/3月切'
    no = 0
    row = 10
    for d in r["data"]:
        print(d["HIN_GAI"])
        no += 1
        row += 1
        sheet['A{}'.format(row)] = no
        sheet['B{}'.format(row)] = d["UKEIRE_DT"]
        sheet['C{}'.format(row)] = d["HIN_GAI"]
        sheet['D{}'.format(row)] = d["HIN_NAME"]
        sheet['E{}'.format(row)] = d["UKEIRE_QTY"]
        sheet['F{}'.format(row)] = d["PrcKoryo"]
        sheet['G{}'.format(row)] = d["Koryo"]
    xls_sv = os.path.dirname(os.path.abspath(__file__))
    xls_sv = os.path.abspath(xls_sv + "\\rfbill_sv.xlsx")

    wb.save(xls_sv)

# ----------------------------------------------------------------
if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        form = cgi.FieldStorage()
        dns = form.getvalue('dns', 'newsdc')
        pallet_no = form.getvalue('pallet_no', '')
        id_no = form.getvalue('id_no', '')
        case_qty = form.getvalue('case_qty', '')
        limit = int(form.getvalue('limit', 0))
        sys.stdout = None
        r = main(dns, pallet_no, id_no ,case_qty , limit)
        sys.stdout = sys.__stdout__
        print_json(r)
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--dt1", help="YYYYMMDD", default="", type=str)
        parser.add_argument("--dt2", help="YYYYMMDD", default="", type=str)
        parser.add_argument("--shimuke", help="01,02", default="", type=str)
        parser.add_argument("--tanto", help="01,02", default="", type=str)
        parser.add_argument("--check", help="1", default="", type=str)
        args= parser.parse_args()
        r = main(args.dns, args.dt1, args.dt2, args.shimuke, args.tanto, args.check)
        rfbill(r)
#        import pprint
#        pprint.pprint(r)
    
