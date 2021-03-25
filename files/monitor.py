# -*- coding: utf-8 -*-

import cgi
import os
import sys
import io
import datetime
import pyodbc
import json
from decimal import Decimal
from datetime import datetime, timedelta

class sdcpos:
    message = ''
    def __init__(self):
#       print('__init__')
        self.dns = 'newsdc'
        self.sql = ''
        self.data = []
        self.results = {}
        self.proc_name = ''
        self.proc_prg = ''
        self.proc_ts = ''

    def __del__(self):
#       print('__del__')
        pass
    
    def __str__(self):
        return 'sdcpos.dns:' + self.dns

    def open(self):
        if 'HTTP_HOST' not in os.environ:
            self.dns = 'newsdc4'
        elif os.environ['HTTP_HOST'] == 'w0':
            self.dns = 'newsdc4'
        self.dns = get_req('dns',self.dns)
        constr = 'DSN=' + self.dns
        self.conn = pyodbc.connect(constr)
        self.cursor = self.conn.cursor()

    def proc(self,name):
        sql = "select * from Proc where Name = '%s'" % name
        p = self.conn.execute(sql).fetchone()
        if p:
            self.results["proc_name"] = p.Name.rstrip()
            self.results["proc_prg"] = p.Prg.rstrip()
            self.results["proc_ts"] = p.Ts.strftime('%Y-%m-%d %H:%M:%S')
            return p.Name
        return ""
    
    def execute(self,sql):
        self.sql = sql
        self.data = []
        try:
            self.cursor.execute(self.sql)
            columns = [column[0] for column in self.cursor.description]
            for c in self.cursor.fetchall():
                for i, v in enumerate(c):
                    if isinstance(v, str):
                        c[i]=v.rstrip()
#                    print(i,v,)
                self.data.append(dict(zip(columns, c)))
        except pyodbc.ProgrammingError as e:
            self.results["error"] = 'pyodbc.ProgrammingError'
            self.results["e.type"] = format(type(e))
            self.results["e.args"] = format(e.args)
            for arg in e.args:
                self.results["e.arg"] = arg     #.decode('UTF-8')
#            str = 'a'
#            print( str.decode('utf-8'))
#            self.results["e"] = e
#            self.results["e"] = "{0}".format(e).decode('utf-8')
        except pyodbc.Error as e:
            self.results["error"] = 'pyodbc.Error '
            self.results["e.type"] = format(type(e))
            self.results["e.args"] = format(e.args)
        except:
            self.results["error"] = 'execute():error'

    def print_response(self):
        self.results["dns"] = self.dns
        self.results["sql"] = self.sql
        self.results["message"] = self.message
        self.results["data"] = self.data
#        print('Content-Type:application/json; charset=UTF-8;')
        print('Content-Type:application/json; charset=UTF-8;\n')
#       print(json.dumps(self.results, ensure_ascii=True, indent=4, sort_keys=False, separators=(',', ': '), cls=DatetimeEncoder))
        print(json.dumps(self.results, default=decimal_default_proc, ensure_ascii=True, indent=4, sort_keys=False, separators=(',', ': ')))
#        print(json.dumps(self.results, ensure_ascii=True, indent=4, sort_keys=False, separators=(',', ': ')))
#        print(json.dumps(self.results, default=decimal_default_proc, ensure_ascii=True, indent=4, sort_keys=False, separators=(',', ': ')))
    
def get_req(nm, v):
    form = cgi.FieldStorage()
    if nm in form:
        v = form[nm].value
    return v

def decimal_default_proc(obj):
#    print(type(obj))
    if isinstance(obj, Decimal):
        return float(obj)
    raise TypeError

class DatetimeEncoder(json.JSONEncoder):
    def default(self, obj):
        print(type(obj))
        if isinstance(obj, str):
            return 'a'
        if isinstance(obj, datetime ):
            return obj.strftime("%Y-%m-%d %H:%M:%S")
        if isinstance(obj, Decimal):
            return float(obj)
        return json.JSONEncoder.default(self, obj)

# ----------------------------------------------------------------
def get_where(w, nm, v1, v2):
    if(v1 == ''):
        return w
    if 'JGYOBU' in nm:
        if v1 == '2,6':
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
def y_syuka_h_cancel():
    sql = '''select
 h.SYUKA_YMD
,h.CANCEL_F
,h.INS_BIN
,h.UNSOU_KAISHA
,h.MUKE_CODE
,h.MUKE_NAME
,h.ID_NO
,h.HIN_NO
,convert(h.SURYO,SQL_DECIMAL) Qty
,convert(h.J_SURYO,SQL_DECIMAL) jQty
,h.KENPIN_NOW
From Y_SYUKA_H h
'''
    w = " where h.SYUKA_YMD = '" + get_req('KAN_DT',datetime.now().strftime('%Y%m%d')) + "'"
    w += " and (h.CANCEL_F = '1'"
    if get_req('UnPick','') != '':
        w += "  or (convert(h.SURYO,SQL_DECIMAL) <> convert(h.J_SURYO,SQL_DECIMAL)"
        w += " and h.MUKE_CODE not in ('7401NF','7401NP','7401NW','7401NZ')"
        w += " and h.MUKE_CODE not in (select distinct TOK_CD from y_syuka_tei))"
    w += " )"
    o = '''
order by
 h.CANCEL_F desc
,h.INS_BIN
,h.MUKE_CODE
'''
    return sql + w + o
# ----------------------------------------------------------------
def y_syuka_h():
    sql = '''select
 h.SYUKA_YMD
,case
 when h.CANCEL_F = '1' then 9
 when h.UNSOU_KAISHA = '佐川急便' then 8
 when h.MUKE_CODE in ('7401NF','7401NP','7401NW','7401NZ') then 7
 when h.UNSOU_KAISHA = '積水' then 6
 when t.TOK_CD is not null then 6
 when h.UNSOU_KAISHA = '' then 5
 else 0
 end
 d_order
,case
 when h.CANCEL_F = '1' then 'キャンセル'
 when t.TOK_CD is not null then '積水'
 when h.MUKE_CODE in ('7401NF','7401NP','7401NW','7401NZ') then 'ミサワ'
 when h.UNSOU_KAISHA = '佐川急便' then '引取り'
 when h.UNSOU_KAISHA = '' then '緊急'
 else rtrim(h.UNSOU_KAISHA)
 end
 UNSOU_KAISHA
,count(distinct LEFT(h.ID_NO,7)) DlvQty
,count(distinct if(INS_BIN = '01',LEFT(ID_NO,7),null)) DlvQty1
,count(distinct if(INS_BIN = '02',LEFT(ID_NO,7),null)) DlvQty2
,count(distinct if(INS_BIN = '03',LEFT(ID_NO,7),null)) DlvQty3
,count(distinct if(INS_BIN = '09',LEFT(ID_NO,7),null)) DlvQty9
,count(*) Cnt
,sum(if(convert(h.SURYO,SQL_DECIMAL) <> convert(h.J_SURYO,SQL_DECIMAL),1,0)) PicZan
,sum(if(h.KENPIN_NOW <> '',0,1)) CntZan
,sum(convert(h.SURYO,SQL_DECIMAL)) Qty
From Y_SYUKA_H h
left outer join y_syuka_tei t
 on (h.SEK_KEN_NO = t.KEN_NO and h.SEK_HIn_NO = t.HIN_NO)
'''
    w = "where h.SYUKA_YMD = '" + get_req('KAN_DT',datetime.now().strftime('%Y%m%d')) + "'"
    
    o = '''
group by
 SYUKA_YMD
,d_order
,UNSOU_KAISHA
order by
 SYUKA_YMD
,d_order
,UNSOU_KAISHA desc
'''
    return sql + w + o
# ----------------------------------------------------------------
def y_syuka():
    y = """
select
 KEY_ID_NO
,'4' Stts
,KEY_SYUKA_YMD
,SYUKA_YMD
,JGYOBA
,HAN_KBN
,CHOKU_KBN
,KEY_MUKE_CODE
,MUKE_NAME
,KAN_KBN
,KENPIN_YMD
,KEY_CYU_KBN
,if(JGYOBA = '00036003',LK_SEQ_NO,'-') LK_SEQ_NO
,jgyobu
,naigai
,key_hin_no
,convert(SURYO,SQL_DECIMAL) Qty
,LK_MUKE_CODE
from y_syuka
where KEY_SYUKA_YMD >= '{0}'
and ((DATA_KBN in ('1','3')) or (DT_SYU = 'R' and KEY_MUKE_CODE <> ''))
//and JGYOBA = '00036003'
union all
select
 IDNo       KEY_ID_NO
,Stts
,SyukaDt    KEY_SYUKA_YMD
,SyukaDt    SYUKA_YMD
,JCode      JGYOBA
,'' HAN_KBN
,'' CHOKU_KBN
,'' KEY_MUKE_CODE
,'' MUKE_NAME
,'' KAN_KBN
,'' KENPIN_YMD
,CyuKbn KEY_CYU_KBN
,'' LK_SEQ_NO
,case KJCode
 when '00023510' then 'D'
 when '00023410' then '4'
 when '00021397' then '5'
 when '00023210' then '7'
 else KJCode
 end jgyobu
,'1'        naigai
,Pn         key_hin_no
,Qty
,''
from HMTAH015
where Stts = '3'
and Soko in (select Value from Config where Name = 'Soko')
""".format(datetime.now().strftime('%Y%m%d'))
    
    sql = """select
 if(y.KEY_SYUKA_YMD <> y.SYUKA_YMD, y.SYUKA_YMD, y.KEY_SYUKA_YMD) KEY_SYUKA_YMD
,y.Stts
,y.JGYOBA
,case
 when y.JGYOBA = '00036003' and y.KEY_MUKE_CODE like 'A%' then '0'
 when y.JGYOBA = '00036003' then '1'
 else '9'
 end ListOrder
,y.HAN_KBN
,y.CHOKU_KBN
,case
 when y.CHOKU_KBN = '1' then ''
 else y.KEY_MUKE_CODE
 end DestCode
,case
 when y.Stts = '3' then '(伝発待ち)'
 when y.HAN_KBN = '2' then '海外' + if(y.KEY_SYUKA_YMD<>y.SYUKA_YMD,' 先行出荷(' + right(y.KEY_SYUKA_YMD,4) + ')','')
// when y.KEY_MUKE_CODE = 'A3' then '東日本サテ'
// when y.KEY_MUKE_CODE = 'A6' then '西日本サテ'
// when y.KEY_MUKE_CODE = 'A7' then '福岡サテ'
// when d.ChoCode is not null then '直送 P産機'
 when y.CHOKU_KBN = '1' and y.LK_MUKE_CODE = '00027768' then '直送 P産機'
 when y.CHOKU_KBN = '1' then '直送 他'
 when ifnull(convert(m.DISPLAY_RANKING,sql_decimal),0) > 0 then m.MUKE_NAME + if(y.jgyobu = 'A',' エアコン','')
 when y.MUKE_NAME = '' then y.KEY_MUKE_CODE
 else y.MUKE_NAME
 end
 Dest
,count(*) Cnt
,sum(if(KAN_KBN = '9',0,1)) CntZan9
,sum(if(KENPIN_YMD <> '',0,1)) CntZan
,sum(if(KEY_CYU_KBN = 'E' or RTrim(LK_SEQ_NO)<>'',0,1)) CntZanLK
,sum(if(KEY_CYU_KBN not in ('1','3'),1,0)) Cnt2
,sum(if(KEY_CYU_KBN not in ('1','3'),if(KAN_KBN = '9',0,1),0)) Cnt2Zan9
,sum(if(KEY_CYU_KBN not in ('1','3'),if(KENPIN_YMD <> '',0,1),0)) Cnt2Zan
,sum(if(KEY_CYU_KBN in ('1','3'),1,0)) Cnt3
,sum(if(KEY_CYU_KBN in ('1','3'),if(KAN_KBN = '9',0,1),0)) Cnt3Zan9
,sum(if(KEY_CYU_KBN in ('1','3'),if(KENPIN_YMD <> '',0,1),0)) Cnt3Zan
,sum(y.Qty) Qty
,sum(if(ifnull(p.JS0,0) > 0,0,ifnull(iSize.Size,0) * y.Qty)) Sai
from ({0}) y
left outer join Item i on (y.jgyobu = i.jgyobu and y.naigai = i.naigai and y.key_hin_no = i.hin_gai)
left outer join ItemSize iSize on (y.jgyobu = iSize.jgyobu and y.key_hin_no = iSize.hin_gai)
left outer join HtDrctId d on (d.IDNo = y.KEY_ID_NO)
left outer join ySize p
on (y.KEY_SYUKA_YMD = p.KEY_SYUKA_YMD
and	y.KEY_MUKE_CODE = p.KEY_MUKE_CODE
and	y.KEY_HIN_NO = p.KEY_HIN_NO)
left outer join Mts m on (m.MUKE_CODE = y.KEY_MUKE_CODE)
""".format(y)
    
    sql += """
group by
 KEY_SYUKA_YMD
,Stts
,ListOrder
,y.JGYOBA
,y.HAN_KBN
,y.CHOKU_KBN
,DestCode
,Dest
order by
 KEY_SYUKA_YMD
,Stts desc
,ListOrder
,y.JGYOBA
,y.HAN_KBN
,y.CHOKU_KBN
,DestCode
,Dest
"""
    return sql
# ----------------------------------------------------------------
def order():
    sql = '''select
 if(y.SYUKA_YMD <> y.KEY_SYUKA_YMD,y.SYUKA_YMD,y.KEY_SYUKA_YMD) KEY_SYUKA_YMD
,y.KEY_SYUKA_YMD KEY_SYUKA_YMD_0
,KEY_MUKE_CODE
,LK_MUKE_CODE
,MUKE_NAME
,CYU_KBN
,CYU_KBN_NAME
,ODER_NO
,count(*) cnt
,sum(if(KAN_KBN='9',0,1)) zan0
,sum(if(KENPIN_YMD<>'',0,1)) zan9
,max(KAN_YMD) zan0YMD
,max(KENPIN_YMD) zan9YMD
,sum(convert(SURYO,sql_decimal)) qty
from Y_Syuka y
'''
    w = "where KEY_CYU_KBN = 'E'"
    o = '''
group by
 KEY_SYUKA_YMD
,KEY_SYUKA_YMD_0
,KEY_MUKE_CODE
,LK_MUKE_CODE
,MUKE_NAME
,CYU_KBN
,CYU_KBN_NAME
,ODER_NO
order by
 KEY_SYUKA_YMD
,KEY_SYUKA_YMD_0
,KEY_MUKE_CODE
,LK_MUKE_CODE
,MUKE_NAME
,CYU_KBN
,CYU_KBN_NAME
,ODER_NO
'''
    return sql + w + o
# ----------------------------------------------------------------
def p_sshiji_splan():
    kan_dt = get_req('KAN_DT',datetime.now().strftime('%Y%m%d'))
    ac_noki = kan_dt[0:4] + '-' + kan_dt[4:6] + '-' + kan_dt[6:8]
    sql = '''
select
distinct
 y.SHIJI_NO SHIJI_NO
,y.YOTEI_DT YOTEI_DT
,y.AcNoki
,case
 when y.AcNoki = '{3}' then '2'                             //当日出荷分
 when y.AcNoki <> '' and ifnull(z.qty92,0) > 0 then '1'     //出荷分
 when y.AcNoki <> '' and ifnull(s.sumiQty,0) > 0 then '1'   //出荷分 済
 when ifnull(s.sumiQty,0) > 0 then ifnull(o.SHIJI_F,'')
 when o.SHIJI_F = '0' and ifnull(z.qty92,0) <= 0 then ''
 when ifnull(o.SHIJI_F,'') = '' and y.AcRow > 0 then ''
 when ifnull(o.SHIJI_F,'') = '' and y.YOTEI_DT <= '{0}' then '0'
 else ifnull(o.SHIJI_F,'')
 end SHIJI_F
,ifnull(o.CANCEL_F,'') CANCEL_F
,ifnull(o.HAKKO_DT,'') HAKKO_DT
,ifnull(o.PRINT_DATETIME,'') PRINT_DATETIME
,y.JGYOBU
,y.HIN_GAI
,'' HIN_NAME
,ifnull(o.KAN_F,'') KAN_F
,ifnull(o.KAN_DT,'') KAN_DT
,ifnull(o.SHIMUKE_CODE,'') SHIMUKE_CODE
,ifnull(o.UKEHARAI_CODE,'') UKEHARAI_CODE
,ifnull(o.TORI_KBN,'') TORI_KBN
,ifnull(o.HIN_CHECK_TANTO,'') HIN_CHECK_TANTO
,if(o.SHIJI_QTY is not null,convert(o.SHIJI_QTY,sql_decimal), y.YOTEI_QTY) qty
,ifnull(z.qty,0) zqty
,ifnull(z.qty92,0) zqty92
,ifnull(z.qty92today,0) zqty92today
,ifnull(z.qtySumi,0) zqtySumi
,ifnull(z.qtyMi,0) zqtyMi
,ifnull(z.GOODS_YMD,'') GOODS_YMD
,ifnull(s.sumiQty,0) sumiQty92
,ifnull(s.miQty,0) miQty92
from SPlan y
left outer join p_sshiji_o o on (y.SHIJI_NO = o.SHIJI_NO)
left outer join (
    select
    JGYOBU
    ,NAIGAI
    ,HIN_GAI
    ,sum(convert(YUKO_Z_QTY,sql_decimal)) qty
    ,sum(if(Soko_No = '92' and NYUKO_DT < '{1}',convert(YUKO_Z_QTY,sql_decimal),0)) qty92
    ,sum(if(Soko_No = '92' and NYUKO_DT >= '{1}',convert(YUKO_Z_QTY,sql_decimal),0)) qty92today
    ,sum(if(Soko_No <> '92' and GOODS_ON =  '0',convert(YUKO_Z_QTY,sql_decimal),0)) qtySumi
    ,sum(if(Soko_No <> '92' and GOODS_ON <> '0',convert(YUKO_Z_QTY,sql_decimal),0)) qtyMi
    ,max(GOODS_YMD) GOODS_YMD
    from zaiko
    group by JGYOBU ,NAIGAI ,HIN_GAI
    ) z on (z.jgyobu = y.jgyobu and z.naigai = '1' and z.hin_gai = y.hin_gai)
left outer join (
    select
     HIN_GAI
    ,sum(convert(SUMI_JITU_QTY,sql_decimal)) sumiQty
    ,sum(convert(MI_JITU_QTY,sql_decimal)) miQty
    from p_sagyo_log
    where jitu_dt = '{1}'
    and FROM_SOKO = '92'
    group by
     HIN_GAI
    ) s on (s.hin_gai = y.hin_gai)
where qty > 0
and CANCEL_F <> '1'
and (KAN_F <> '1' or KAN_DT = '{2}')
and TORI_KBN in ('3','')
and SHIMUKE_CODE in ('01','02','03')
and (zqty92 > 0 or zqty92today > 0 or sumiQty92 > 0 or miQty92 > 0)
order by
 SHIJI_F desc
,y.YOTEI_DT
,if(y.AcNoki = '','1','0')
,y.AcNoki
,o.HAKKO_DT
,y.SHIJI_NO
'''
    tday = datetime.now().strftime('%Y%m%d')
    return sql.format( kan_dt, tday, tday, ac_noki)
# ----------------------------------------------------------------
def p_sshiji_plan():
    kan_dt = get_req('KAN_DT',datetime.now().strftime('%Y%m%d'))
    sql = '''select
 o.SHIJI_NO
,o.YOTEI_DT
,o.SHIJI_F
,o.CANCEL_F
,o.HAKKO_DT
,o.PRINT_DATETIME
,o.JGYOBU
,o.HIN_GAI
,'' HIN_NAME
,o.KAN_F
,o.KAN_DT
,o.UKEHARAI_CODE
,o.HIN_CHECK_TANTO
,o.qty
,ifnull(z.qty,0) zqty
,ifnull(z.qty92,0) zqty92
,ifnull(z.GOODS_YMD,'') GOODS_YMD
,ifnull(s.sumiQty,0) sumiQty92
,ifnull(s.miQty,0) miQty92
from (
    select
     y.KEY_NO SHIJI_NO
    ,y.YOTEI_DT
    ,'0' SHIJI_F
    ,'' CANCEL_F
    ,left(SASIZU_DateTime,8) HAKKO_DT
    ,SASIZU_DateTime PRINT_DATETIME
    ,y.JGYOBU
    ,y.NAIGAI
    ,y.HIN_GAI
    ,if(y.S_KAN_DateTime = '','0','1') KAN_F
    ,left(y.S_KAN_DateTime,8) KAN_DT
    ,left(y.TEHAISAKI,3) UKEHARAI_CODE
    ,'' HIN_CHECK_TANTO
    ,convert(y.YOTEI_QTY,sql_decimal) qty
    from PLN_S_YOTEI y
) o
left outer join (
    select
    JGYOBU
    ,NAIGAI
    ,HIN_GAI
    ,sum(convert(YUKO_Z_QTY,sql_decimal)) qty
    ,sum(if(Soko_No = '92',convert(YUKO_Z_QTY,sql_decimal),0)) qty92
    ,max(GOODS_YMD) GOODS_YMD
    from zaiko
    group by JGYOBU ,NAIGAI ,HIN_GAI
    ) z on (z.jgyobu = o.jgyobu and z.naigai = o.naigai and z.hin_gai = o.hin_gai)
left outer join (
    select
     HIN_GAI
    ,sum(convert(SUMI_JITU_QTY,sql_decimal)) sumiQty
    ,sum(convert(MI_JITU_QTY,sql_decimal)) miQty
    from p_sagyo_log
'''
    sql += "where jitu_dt = '" + datetime.now().strftime('%Y%m%d') + "'"
    sql += '''
    and FROM_SOKO = '92'
    group by
     HIN_GAI
    ) s on (s.hin_gai = o.hin_gai)
where o.qty > 0
order by
 YOTEI_DT
,SHIJI_F desc
,SHIJI_NO
'''
    return sql
# ----------------------------------------------------------------
def p_sshiji():
    sql = '''select
 o.SHIJI_NO
,o.SHIJI_F
,o.CANCEL_F
,o.HAKKO_DT
,o.PRINT_DATETIME
,o.JGYOBU
,o.HIN_GAI
,i.HIN_NAME
,o.KAN_F
,o.KAN_DT
,o.UKEHARAI_CODE
,o.HIN_CHECK_TANTO
,convert(o.SHIJI_QTY,sql_decimal) qty
,ifnull(z.qty,0) zqty
,ifnull(z.qty92,0) zqty92
,ifnull(z.qtySumi,0) zqtySumi
,ifnull(z.qtyMi,0) zqtyMi
,ifnull(z.GOODS_YMD,'') GOODS_YMD
,ifnull(s.sumiQty,0) sumiQty92
,ifnull(s.miQty,0) miQty92
from p_sshiji_o o
left outer join item i on (i.jgyobu = o.jgyobu and i.naigai = o.naigai and i.hin_gai = o.hin_gai)
left outer join (
    select
    JGYOBU
    ,NAIGAI
    ,HIN_GAI
    ,sum(convert(YUKO_Z_QTY,sql_decimal)) qty
    ,sum(if(Soko_No = '92',convert(YUKO_Z_QTY,sql_decimal),0)) qty92
    ,sum(if(Soko_No <> '92' and GOODS_ON =  '0',convert(YUKO_Z_QTY,sql_decimal),0)) qtySumi
    ,sum(if(Soko_No <> '92' and GOODS_ON <> '0',convert(YUKO_Z_QTY,sql_decimal),0)) qtyMi
    ,max(GOODS_YMD) GOODS_YMD
    from zaiko
    group by JGYOBU ,NAIGAI ,HIN_GAI
    ) z on (z.jgyobu = o.jgyobu and z.naigai = o.naigai and z.hin_gai = o.hin_gai)
left outer join (
    select
     HIN_GAI
    ,sum(convert(SUMI_JITU_QTY,sql_decimal)) sumiQty
    ,sum(convert(MI_JITU_QTY,sql_decimal)) miQty
    from p_sagyo_log
'''
    sql += "where jitu_dt = '" + datetime.now().strftime('%Y%m%d') + "'"
    sql += '''
    and FROM_SOKO = '92'
    group by
     HIN_GAI
    ) s on (s.hin_gai = o.hin_gai)
'''
    w = '''
where o.CANCEL_F <> '1'
and ((o.HAKKO_DT >= '20171001' and o.KAN_DT = '20180310')
or (o.HAKKO_DT >= '20171225' and o.KAN_DT = '' and o.SHIJI_F <> '0')
or (o.HAKKO_DT >= '20171225' and o.KAN_DT = '' and o.SHIJI_F = '0' and zqty92 > 0) )
'''
#    tstr = datetime.now().strftime('%Y%m%d')
    w = "where o.CANCEL_F <> '1'"
    w = get_where(w,'o.UKEHARAI_CODE',get_req('UKEHARAI_CODE',''),'')
#    w += " and ((o.HAKKO_DT >= '20171001' and o.KAN_DT = '" + datetime.now().strftime('%Y%m%d') + "')"
    w += " and ((o.KAN_DT = '" + get_req('KAN_DT',datetime.now().strftime('%Y%m%d')) + "')"
    hakko_dt = get_req('HAKKO_DT','20171225')
    w += " or (o.HAKKO_DT >= '" + hakko_dt + "' and o.KAN_DT = '' and o.SHIJI_F <> '0')"
    if get_req('UKEHARAI_CODE','') == 'ZN8':
        w += " or (o.HAKKO_DT >= '" + hakko_dt + "' and o.KAN_DT = '' and o.SHIJI_F = '0') )"
    elif get_req('UKEHARAI_CODE','') == 'ZG7':
        hakko_dt = get_req('HAKKO_DT','20180301')
        w += " or (o.HAKKO_DT >= '" + hakko_dt + "' and o.KAN_DT = '' and o.SHIJI_F = '0') )"
    elif get_req('UKEHARAI_CODE','') == 'ZH0':
        hakko_dt = get_req('HAKKO_DT','20180301')
        w += " or (o.HAKKO_DT >= '" + hakko_dt + "' and o.KAN_DT = '' and o.SHIJI_F = '0') )"
    elif get_req('UKEHARAI_CODE','') == 'ZF0':
        hakko_dt = get_req('HAKKO_DT','20180515')
        w += " or (o.HAKKO_DT >= '" + hakko_dt + "' and o.KAN_DT = '' and o.SHIJI_F = '0') )"
    else:
        w += " or (o.HAKKO_DT >= '" + hakko_dt + "' and o.KAN_DT = '' and o.SHIJI_F = '0' and (zqty92 > 0 or sumiQty92 > 0 or miQty92 > 0)))"
    o = '''
order by
 SHIJI_F desc
,o.SHIJI_NO
'''
    return sql + w + o
# ----------------------------------------------------------------
def p_sshiji_92():
    sql = '''select
 '' SHIJI_NO
,'0' SHIJI_F
,'' CANCEL_F
,'' HAKKO_DT
,'' PRINT_DATETIME
,z.JGYOBU
,z.HIN_GAI
,i.HIN_NAME
,if(z.qty92 = 0,'1','0') KAN_F
,'' KAN_DT
,'' UKEHARAI_CODE
,'' HIN_CHECK_TANTO
,z.qty
,z.qty zqty
,z.qty92 zqty92
,'' GOODS_YMD
,0 sumiQty92
,0 miQty92
from (
    select
    JGYOBU
    ,NAIGAI
    ,HIN_GAI
    ,sum(convert(YUKO_Z_QTY,sql_decimal)) qty
    ,sum(if(Soko_No = '92',convert(YUKO_Z_QTY,sql_decimal),0)) qty92
    ,sum(if(Soko_No = '95',convert(YUKO_Z_QTY,sql_decimal),0)) qty95
    from zaiko
    where JGYOBU = 'A' and Soko_No in ('92','95')
    group by JGYOBU ,NAIGAI ,HIN_GAI
) z
left outer join item i on (i.jgyobu = z.jgyobu and i.naigai = z.naigai and i.hin_gai = z.hin_gai)
'''
    w = ""
    o = '''
'''
    return sql + w + o
# ----------------------------------------------------------------
def p_shorder():
    sql = '''select
 s.ORDER_NO
,s.ORDER_DT
,s.TANTO_CODE
,s.JGYOBU
,s.NAIGAI
,s.HIN_GAI
,s.ORDER_CODE
,s.DELI_CODE
,convert(s.ORDER_QTY,sql_decimal) qty
,if(s.ANS_NOUKI_DT <> '',s.ANS_NOUKI_DT,s.Y_NOUKI_DT)
 Y_NOUKI_DT
,s.TANKA
,s.LOT
,s.KAN_F
,ifnull(p.NYUKA_DT,s.KAN_DT) KAN_DT
,s.BUNNOU_CNT
,s.UKEIRE_QTY
,s.CANCEL_F
,s.CANCEL_DATETIME
,s.PRINT_F
,s.WS_NO
,s.G_SHIIRE_KBN
,s.G_SYUSHI
,s.TORI_KBN
,s.ANS_NOUKI_DT
,s.USE_YM
,p.cnt pCnt
,p.qty pQty
,p.NYUKA_DT pDt
,ifnull(z.qty,0) zQty
from P_SHORDER s
left outer join (
    select
     JGYOBU
    ,NAIGAI
    ,HIN_GAI
    ,count(*) cnt
    ,sum(convert(NYUKA_QTY,sql_decimal)) qty
    ,max(NYUKA_DT) NYUKA_DT
    from p_nyuka
    group by
     JGYOBU
    ,NAIGAI
    ,HIN_GAI
) p on (s.JGYOBU = p.JGYOBU and s.NAIGAI = p.NAIGAI and s.HIN_GAI = p.HIN_GAI)
left outer join (
    select
    JGYOBU
    ,NAIGAI
    ,HIN_GAI
    ,sum(convert(YUKO_Z_QTY,sql_decimal)) qty
    from zaiko
    group by JGYOBU ,NAIGAI ,HIN_GAI
) z on (s.JGYOBU = z.JGYOBU and s.NAIGAI = z.NAIGAI and s.HIN_GAI = z.HIN_GAI)
'''
    w = "where ((s.CANCEL_F <> '1' and s.KAN_F <> '1')"
    w += " or (s.KAN_DT >= '%s'))" % datetime.now().strftime('%Y%m%d')
    jgyobu = get_req('JGYOBU','');
    if jgyobu == 'A':
        w += " and s.G_SYUSHI = '120'"
    elif jgyobu == 'R':
        w += " and s.G_SYUSHI = '220'"
    o = '''
order by
 KAN_DT
,Y_NOUKI_DT
,s.ORDER_CODE
'''
    return sql + w + o
#############
#p_sagyo_log
#############
def p_sagyo_log():
    sql = '''select
 p.TANTO_CODE
,case t.TANTO_NAME
 when '田中馨' then '田中馨'
 when '田中みちよ' then '田中み'
 when '柳昌美' then '柳'
 when '長谷川君代' then '長谷川'
 when '谷紹未' then '谷'
 when '松井まいむ' then '松井ま'
 when '井ノ口泰司' then '井ノ口'
 when '久木田なおみ' then '久木田'
 when '林秀雄' then '林'
 when '林良子' then '林'
 when '佐々木奈々' then '佐々木'
 else left(t.TANTO_NAME,LENGTH('ああ'))
 end
 TANTO_NAME
,p.JITU_DT
,p.JITU_TM
,p.RIRK_ID
,p.FROM_SOKO
,p.TO_SOKO
,if(ifnull(sfr.SOKO_BUN,'9') < ifnull(sto.SOKO_BUN,'9')
   ,ifnull(sfr.SOKO_BUN,'9'),ifnull(sto.SOKO_BUN,'9'))
 SOKO_BUN
,m.MENU_DSP
,case
 when sfr.SOKO_BUN = '0' then if(p.FROM_SOKO = left(sfr.SOKO_NAME,2),sfr.SOKO_NAME,p.FROM_SOKO + ' ' + sfr.SOKO_NAME)
 when sto.SOKO_BUN = '0' then if(p.TO_SOKO = left(sto.SOKO_NAME,2),sto.SOKO_NAME,p.TO_SOKO + ' ' + sto.SOKO_NAME)
 else m.MENU_DSP
 end Loc
from p_sagyo_log p
left outer join tanto t on (p.TANTO_CODE = t.TANTO_CODE)
left outer join soko sfr on (p.FROM_SOKO = sfr.Soko_No)
left outer join soko sto on (p.TO_SOKO = sto.Soko_No)
left outer join P_MENU m
 on (m.JGYOBU = p.JGYOBU and m.NAIGAI = p.NAIGAI and m.MENU_NO = p.MENU_NO)
'''
    now = datetime.now()
    jitu_dt = get_req('JITU_DT','')
    if(jitu_dt == ''):
        jitu_dt = now.strftime('%Y%m%d')
        w = "where p.JITU_DT = '" + jitu_dt + "'"
        now -= timedelta(minutes=30)
        jitu_tm = now.strftime('%H%M%S')
        w += " and p.JITU_TM > '" + jitu_tm + "'"
    else:
        w = "where p.JITU_DT = '" + jitu_dt + "'"
        jitu_tm1 = now.strftime('%H%M%S')
        now -= timedelta(minutes=60)
        jitu_tm2 = now.strftime('%H%M%S')
        w += " and p.JITU_TM between '" + jitu_tm2 + "' and '" + jitu_tm1 + "'"
    w += " and Loc <> ''"
    w = get_where(w,'p.JGYOBU',get_req('JGYOBU',''),'')
    o = '''
order by
 p.TANTO_CODE
,SOKO_BUN
,p.JITU_DT desc
,p.JITU_TM desc
'''
    return sql + w + o
#############
#AcOrder
#############
def AcOrder():
    sql = '''
select
top 3
 Noki NokiOrg
,case
 when Noki = '' then '空白'
 when Noki like '%-%-%' then right(rtrim(Noki),5)
 when Noki like '%/%/%' then right(rtrim(Noki),5)
 else Noki
 End NokiDsp
,count(
	if(KanDt like '%-%-%' or KanDt in ('','SX','ＳＸ'),KanDt,null))	sdcCnt
,sum(
	if(KanDt like '%-%-%' or KanDt in ('','SX','ＳＸ'),Qty,0))		sdcQty
,count(
	if(KanDt = '',KanDt,null))	zanCnt
,sum(
	if(KanDt = '',Qty,0))	zanQty
,count(
	if(KanDt like '%-%-%' or KanDt in ('','SX','ＳＸ'),null,KanDt))	othCnt
,sum(
	if(KanDt like '%-%-%' or KanDt in ('','SX','ＳＸ'),0,Qty))		othQty
,count(*) 	Cnt
,sum(Qty)	Qty
from AcOrder
'''
    now = datetime.now()
    w = "where Qty > 0"
    w += " and Noki >= '" + now.strftime('%Y-%m-%d') + "'"
    g = '''
group by
 Noki
order by
 NokiOrg
,NokiDsp
'''
    return sql + w + g
#############
#zaiko9
#############
def zaiko9():
    sql = '''
select
 z.Soko
,s.SOKO_NAME
,z.Tana
,z.HIN_GAI Pn
,z.Qty
,z.inDate
,i.ST_SOKO + i.ST_RETU + i.ST_REN + i.ST_DAN stTana
from (
    select
     Soko_No Soko
    ,Soko_No + Retu + Ren + Dan Tana
    ,jgyobu
    ,naigai
    ,HIN_GAI
    ,sum(convert(YUKO_Z_QTY,sql_decimal)) Qty
    ,max(NYUKO_DT) inDate
    from zaiko
    where Soko_No in ('90','94','95','96','98')
    group by
     Soko
    ,Tana
    ,jgyobu
    ,naigai
    ,HIN_GAI
) z
left outer join Soko s on (s.Soko_No = z.Soko)
left outer join item i on (i.jgyobu = z.jgyobu and i.naigai = z.naigai and i.hin_gai = z.hin_gai)
'''
    w = get_where('','z.JGYOBU',get_req('JGYOBU',''),'')
    o = '''
order by
 z.Tana
,stTana
,z.inDate
'''
    return sql + w + o
#############
#メイン
#############
if __name__ == "__main__":
    s = sdcpos()
#    print(s)
    s.open()
    proc = ""
    sql = "select DATABASE(),USER(),'aaa    '"
#    sql = "SELECT * FROM dbo.fSQLTables(null, null, null)"
    f = get_req('f','')
    if f  == 'p_sshiji':
        if get_req('UKEHARAI_CODE','') == 'PLAN':
            proc = "SPlan"
            if get_req('splan','') == '':
                import splan
                sys.stdout = open('nul', 'w')
                sp = splan.splan()
                sp.main(s.dns)
                sys.stdout = sys.__stdout__
            sql = p_sshiji_splan()
        elif get_req('UKEHARAI_CODE','') == 'PLAN.':
            proc = "SPlan"
            sql = p_sshiji_splan()
        elif get_req('UKEHARAI_CODE','') == 'AC':
            proc = "SPlan"
            if get_req('splan','') == '':
                import splan
                sys.stdout = open('nul', 'w')
                sp = splan.splan()
                sp.main_ac(s.dns)
                sys.stdout = sys.__stdout__
            sql = p_sshiji_splan()
        elif get_req('UKEHARAI_CODE','') == 'AC.':
            proc = "SPlan"
            sql = p_sshiji_splan()
        else:
            sql = p_sshiji()
        if proc == "SPlan":
            s.message = "・前日在庫92にあるものだけを表示するように変更"
    elif f  == 'p_sshiji_92':
        sql = p_sshiji_92()
    elif f  == 'y_syuka_h':
        sql = y_syuka_h()
    elif f  == 'y_syuka_h_cancel':
        sql = y_syuka_h_cancel()
    elif f  == 'y_syuka':
        sql = y_syuka()
    elif f  == 'order':
        sql = order()
    elif f  == 'p_shorder':
        sql = p_shorder()
    elif f  == 'p_sagyo_log':
        sql = p_sagyo_log()
    elif f  == 'AcOrder':
        sql = AcOrder()
    elif f  == 'zaiko9':
        sql = zaiko9()
#    sql = p_sshiji()
    if s.proc(proc) == "":
        s.execute(sql)
    s.print_response()
#   print(s.dns)
#   del s
