# -*- coding: utf-8 -*-
import cgi
import os
import sys
import io
import pyodbc
import json
from decimal import Decimal
from datetime import datetime, timedelta

_debug = False
def debug(v):
    if _debug:
        print(v)

def decimal_default_proc(obj):
    if isinstance(obj, Decimal):
        return float(obj)
    raise TypeError

def get_req(nm, v):
    form = cgi.FieldStorage()
    if nm in form:
        v = form[nm].value
    return v

class p_sshiji:
    _dns = "newsdc"
    _JGYOBU = ""
    _HAKKO_DT = "20180101"
    _KAN_DT = format(datetime.now().strftime('%Y%m%d'))
    _UKEHARAI_CODE = ""
    TORI_KBN = "3"

    @property
    def dns(self):
        return self._dns
    @dns.setter
    def dns(self, value):
        self._dns = value
    @property
    def JGYOBU(self):
        return self._JGYOBU.upper()
    @JGYOBU.setter
    def JGYOBU(self, value):
        self._JGYOBU = value
    @property
    def UKEHARAI_CODE(self):
        return self._UKEHARAI_CODE.upper()
    @UKEHARAI_CODE.setter
    def UKEHARAI_CODE(self, value):
        self._UKEHARAI_CODE = value
        if value == 'ZN8':
            self.TORI_KBN = ''
    @property
    def HAKKO_DT(self):
        if self.UKEHARAI_CODE == "ZF0":
            return "20180801"
        if self.UKEHARAI_CODE == "ZN8":
            return "20181001"
        return self._HAKKO_DT
    @HAKKO_DT.setter
    def HAKKO_DT(self, value):
        self._HAKKO_DT = value
    @property
    def KAN_DT(self):
        return self._KAN_DT
    @KAN_DT.setter
    def KAN_DT(self, value):
        self._KAN_DT = value

    def __init__(self):
        pass

    def __del__(self):
        pass
    
    def __str__(self):
        return self.__class__.__name__ + str(self.__dict__)

    def main(self):
        debug(self.__class__.__name__ + '.' + sys._getframe().f_code.co_name)
        debug("connect() DSN=" + self.dns)
        self.conn = pyodbc.connect('DSN=' + self.dns)
        self.list()
        debug("close()")
        self.conn.close()

    def list(self):
        debug(self.__class__.__name__ + '.' + sys._getframe().f_code.co_name)
        data = []
        debug(self.__class__.__name__ + '.' + sys._getframe().f_code.co_name)
        sql = self.get_sql()
        debug(sql)
        cursor = self.conn.cursor()
        cursor.execute(sql)
        columns = [column[0] for column in cursor.description]
        debug(columns)
        for row in cursor.fetchall():
            for i, v in enumerate(row):
                if v is None:
                    row[i]=""
                elif isinstance(v, str):
                    row[i]=v.rstrip()
            dic = dict(zip(columns, row))
            if dic["SHIJI_F"] == "0":
                dic["SHIJI_NAME"] = "事前"
            elif dic["SHIJI_F"] == "1":
                dic["SHIJI_NAME"] = "スポット"
            elif dic["SHIJI_F"] == "2":
                dic["SHIJI_NAME"] = "欠品解除"
            elif dic["SHIJI_F"] == "3":
                dic["SHIJI_NAME"] = "再梱包"
            else:
                dic["SHIJI_NAME"] = dic["SHIJI_F"]
#           debug(dic)
            data.append(dic)
        ### 商品化計画追加
        sql = self.get_sql_plan()
        if sql != "":
            debug(sql)
            cursor = self.conn.cursor()
            cursor.execute(sql)
            columns = [column[0] for column in cursor.description]
            debug(columns)
            for row in cursor.fetchall():
                for i, v in enumerate(row):
                    if isinstance(v, str):
                        row[i]=v.rstrip()
                dic = dict(zip(columns, row))
                debug(dic)
                ### dataにあるかチェック
                for d in data:
                    if d["JGYOBU"] == dic["JGYOBU"] \
                    and d["HIN_GAI"] == dic["HIN_GAI"]:
                        d["YOTEI_DT"] = dic["YOTEI_DT"]
                        dic["YOTEI_DT"] = ""
                        break
                if dic["YOTEI_DT"] != "" and False:
                    ###完了チェック
                    sql = "select top 1 * from P_SSHIJI_O"
                    sql += " where JGYOBU='{0}'".format(dic["JGYOBU"])
                    sql += " and NAIGAI='1'"
                    sql += " and HIN_GAI='{0}'".format(dic["HIN_GAI"])
                    sql += " and CANCEL_F='0'"
                    sql += " and HAKKO_DT >= '{0}'".format(row.Ins_DateTime[:8])
                    sql += " and PRINT_DATETIME > '{0}'".format(row.Ins_DateTime)
                    p_sshiji = self.conn.execute(sql).fetchone()
                    if p_sshiji:
                        dic["SHIJI_NO"] = p_sshiji.SHIJI_NO.rstrip()
                        dic["HAKKO_DT"] = p_sshiji.HAKKO_DT.rstrip()
                        dic["UKEHARAI_CODE"] = p_sshiji.UKEHARAI_CODE.rstrip()
                        dic["KAN_F"] = p_sshiji.KAN_F.rstrip()
                        dic["KAN_DT"] = p_sshiji.KAN_DT.rstrip()
                        debug("SHIJI_NO={0}".format(dic["SHIJI_NO"]))
                    if dic["KAN_DT"] == "":
                        data.append(dic)

        self.print_response(data)

    def print_response(self, data):
        results = {}
        results["dns"] = self.dns
        results["JGYOBU"] = self.JGYOBU
        results["UKEHARAI_CODE"] = self.UKEHARAI_CODE
        results["HAKKO_DT"] = self.HAKKO_DT
        results["KAN_DT"] = self.KAN_DT
        results["data"] = data
        if self.JGYOBU == "A":
            results["message"] = "エアコン スポット→出荷分、欠品解除→当日出荷分"
#            results["message"] = "エアコン 当日92出庫で商品化完了した場合は予定に上げる"
#            results["message"] = "エアコン 当日13:00までに在庫92へ上げた分を予定に含める"
#            results["message"] = "エアコン商品化完了分を検索しないように変更"
        print('Content-Type:application/json; charset=UTF-8;\n')
        print(json.dumps(results, default=decimal_default_proc, ensure_ascii=True, indent=4, sort_keys=False, separators=(',', ': ')))
        debug("len(data)={0}".format(len(data)))

    def get_sql_plan(self):
        # 商品化計画SQL
        debug(self.__class__.__name__ + '.' + sys._getframe().f_code.co_name)
        if self.UKEHARAI_CODE == 'ZI0':
            # 広島
            debug("広島(ZI0)計画参照なし")
            return ""
        if self.UKEHARAI_CODE == 'ZE0':
            debug("滋賀PC(ZE0)計画参照なし")
            return ""
        if self.UKEHARAI_CODE == 'ZB0':
            debug("小野PC(ZB0)計画参照なし")
            return ""

        sql = """select
 p.YOTEI_DT
,p.Ins_DateTime
,'' SHIJI_NO
,'' SHIJI_F
,'' CANCEL_F
,'' HAKKO_DT
,'' PRINT_DATETIME
,p.JGYOBU
,p.HIN_GAI
,'' HIN_NAME
,'' KAN_F
,'' KAN_DT
,p.TEHAISAKI UKEHARAI_CODE
,'' HIN_CHECK_TANTO
,convert(p.YOTEI_QTY,sql_decimal) qty
,0 zqty
,0 zqty92
,0 zqtySumi
,0 zqtyMi
,'' GOODS_YMD
,0 sumiQty92
,0 miQty92
from PLN_S_YOTEI p
where qty > 0
"""
        if self.JGYOBU != "":
            sql += " and p.JGYOBU='{0}'".format(self.JGYOBU)
        sql += """
order by
 p.YOTEI_DT
"""
        return sql

    def get_sql(self):
        debug(self.__class__.__name__ + '.' + sys._getframe().f_code.co_name)
        sql = ""
        if self.UKEHARAI_CODE == 'ZI0':
            debug("広島(ZI0-JCS)")
            sql += """select distinct
 o.ORDER_DT AcNoki
,o.SHIJI_F
,o.SHIJI_NO
,o.CANCEL_F
,o.HAKKO_DT
,o.PRINT_DATETIME
,o.JGYOBU
,o.HIN_GAI
,i.NameJ HIN_NAME
,i.SType
,ifnull(c.SCat,'小物') SCat
,o.KAN_F
,o.KAN_DT
,o.UKEHARAI_CODE
,o.HIN_CHECK_TANTO
,o.HIN_CHECK_DATETIME
,convert(o.SHIJI_QTY,sql_decimal) qty
,0 zqty
,0 zqty92
,0 zqty92today
,0 zqtySumi
,0 zqtyMi
,'' GOODS_YMD
,0 sumiQty92
,0 miQty92
from p_sshiji_o o
left outer join JcsItem i on (o.hin_gai = i.MazdaPn)
left outer join JcsCat c on (i.SType = c.SType)
where o.ORDER_DT >= '{0}'
and o.CANCEL_F <> '1'
union
select distinct
 left(o.DlvDt,4) + SUBSTRING(o.DlvDt,6,2) + right(o.DlvDt,2) AcNoki
,rtrim(o.NG) SHIJI_F
,rtrim(o.NohinNo) + rtrim(o.NohinNo2) SHIJI_NO
,'' CANCEL_F
,replace(o.OrderDt,'/','') HAKKO_DT
,'' PRINT_DATETIME
,'J' JGYOBU
,o.MazdaPn HIN_GAI
,i.NameJ HIN_NAME
,i.SType
,ifnull(c.SCat,'小物') SCat
,'' KAN_F
,'' KAN_DT
,'' UKEHARAI_CODE
,'' HIN_CHECK_TANTO
,'' HIN_CHECK_DATETIME
,convert(o.Qty,sql_decimal) qty
,0 zqty
,0 zqty92
,0 zqty92today
,0 zqtySumi
,0 zqtyMi
,'' GOODS_YMD
,0 sumiQty92
,0 miQty92
from JcsOrder o
left outer join JcsItem i on (o.MazdaPn = i.MazdaPn)
left outer join JcsCat c on (i.SType = c.SType)
where o.NG = ''
""".format(datetime.now().strftime('%Y%m%d'))
            # 広島(ZI0)
            return sql
        if self.UKEHARAI_CODE == 'ZI0-p_sshiji_o':
            debug("広島(ZI0-p_sshiji_o)")
            sql += """select distinct
 o.ORDER_DT AcNoki
,o.SHIJI_F
,o.SHIJI_NO
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
,o.HIN_CHECK_DATETIME
,convert(o.SHIJI_QTY,sql_decimal) qty
,0 zqty
,0 zqty92
,0 zqty92today
,0 zqtySumi
,0 zqtyMi
,'' GOODS_YMD
,0 sumiQty92
,0 miQty92
from p_sshiji_o o
left outer join item i
on (i.jgyobu = o.jgyobu and i.naigai = o.naigai and i.hin_gai = o.hin_gai)
where o.ORDER_DT >= '{0}'
and o.CANCEL_F <> '1'
""".format(datetime.now().strftime('%Y%m%d'))
            # 広島(ZI0)
            return sql

        if self.JGYOBU == "A":
            sql_shiji_f = """
 case
 when o.SHIJI_F = '2' or a.Noki = '{0}' then '2'
 when (o.SHIJI_F = '1' or a.Noki <> '') and ifnull(z.qty92,0) > 0 then '1'
 when (o.SHIJI_F = '1' or a.Noki <> '') and ifnull(s.sumiQty,0) > 0 then '1'
 when ifnull(z.qty92,0) > 0 then '0'
 when ifnull(s.sumiQty,0) > 0 then '0'
 else ''
 end SHIJI_F
,a.QtyS aQtyS
,a.Noki aNoki
,a.Tanto aTanto		
,a.Biko1 aBiko1
,a.KanDt aKanDt
,a.KanQty aKanQty
,a.NaraDt aNaraDt
""".format(datetime.now().strftime('%Y-%m-%d'))
            sql_acnoki = ",a.Noki AcNoki"
            
        elif self.UKEHARAI_CODE == 'ZN8' \
        or self.UKEHARAI_CODE == 'ZG7' \
        or self.UKEHARAI_CODE == 'ZH0' \
        or self.UKEHARAI_CODE == 'ZE0' \
        or self.UKEHARAI_CODE == 'ZB0' \
        or self.UKEHARAI_CODE == 'ZF0':
            sql_shiji_f = "o.SHIJI_F"
            sql_acnoki = "" #"'' AcNoki"
        else:
            sql_shiji_f = """
 case
 when o.SHIJI_F = '0' and ifnull(z.qty92,0) = 0 and ifnull(s.sumiQty,0) = 0 then ''
 else o.SHIJI_F
 end SHIJI_F
"""
            sql_acnoki = "" #"'' AcNoki"
        sql += """
select distinct
{0}
{1}
,o.SHIJI_NO
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
,o.HIN_CHECK_DATETIME
,convert(o.SHIJI_QTY,sql_decimal) qty
,ifnull(z.qty,0) zqty
,ifnull(z.qty92,0) zqty92
,ifnull(z.qty92today,0) zqty92today
,ifnull(z.qtySumi,0) zqtySumi
,ifnull(z.qtyMi,0) zqtyMi
,ifnull(z.GOODS_YMD,'') GOODS_YMD
,ifnull(s.sumiQty,0) sumiQty92
,ifnull(s.miQty,0) miQty92
,y.Qty yQty
,y.Cnt yCnt
,y.Qty2 yQty2
,if(ifnull(z.qty92,0) > 0 and ifNull(y.Qty,0) > 0 and ifNull(y.Qty,0) >= ifnull(z.qtySumi,0),ifNull(y.Qty,0) - ifnull(z.qtySumi,0),-1) zOrder
from p_sshiji_o o
""".format(sql_shiji_f, sql_acnoki)
        
        sql += """
left outer join item i
on (i.jgyobu = o.jgyobu and i.naigai = o.naigai and i.hin_gai = o.hin_gai)
"""
        sql += """
left outer join (
select
    z.JGYOBU
    ,z.NAIGAI
    ,z.HIN_GAI
    ,sum(convert(z.YUKO_Z_QTY,sql_decimal)) qty
    ,sum(if(z.Soko_No = '92' and (z.NYUKO_DT < '{0}' or ifnull(p.Qty,0) > 0),convert(z.YUKO_Z_QTY,sql_decimal),0)) qty92
    ,sum(if(z.Soko_No = '92' and (z.NYUKO_DT >= '{0}' and ifnull(p.Qty,0) <= 0),convert(z.YUKO_Z_QTY,sql_decimal),0)) qty92today
    ,sum(if(z.Soko_No <> '92' and z.GOODS_ON =  '0',convert(z.YUKO_Z_QTY,sql_decimal),0)) qtySumi
    ,sum(if(z.Soko_No <> '92' and z.GOODS_ON <> '0',convert(z.YUKO_Z_QTY,sql_decimal),0)) qtyMi
    ,max(z.GOODS_YMD) GOODS_YMD
    from zaiko z
	left outer join (
select
 JGYOBU
,NAIGAI
,HIN_GAI
,sum(convert(SUMI_JITU_QTY,sql_decimal)+convert(MI_JITU_QTY,sql_decimal)) Qty
from p_sagyo_log
where jitu_dt = '{0}'
and jitu_tm < '130000'
and TO_SOKO = '92'
group by
 JGYOBU
,NAIGAI
,HIN_GAI
	) p
		on (z.JGYOBU = p.JGYOBU and z.NAIGAI=p.NAIGAI and z.HIN_GAI=p.HIN_GAI)
    group by z.JGYOBU ,z.NAIGAI ,z.HIN_GAI
    ) z on (z.jgyobu = o.jgyobu and z.naigai = o.naigai and z.hin_gai = o.hin_gai)
""".format(datetime.now().strftime('%Y%m%d'))
        sql += """
left outer join (
    select
     HIN_GAI
    ,sum(convert(SUMI_JITU_QTY,sql_decimal)) sumiQty
    ,sum(convert(MI_JITU_QTY,sql_decimal)) miQty
    from p_sagyo_log
    where jitu_dt = '{0}'
    and FROM_SOKO = '92'
    group by
     HIN_GAI
    ) s on (s.hin_gai = o.hin_gai)
""".format(datetime.now().strftime('%Y%m%d'))
        #出荷予定
        sql += """
left outer join (
select
 Pn
,sum(Qty) Qty
,count(*) Cnt
,sum(if(Stts='2',Qty,0)) Qty2
from HMTAH015
where Soko in (select Value from Config where Name = 'Soko')
and IDNo not in (select distinct KEY_ID_NO from y_syuka where KAN_KBN = '9')
group by
 Pn
) y on (y.Pn = o.hin_gai)
"""
        if self.JGYOBU == "A":
            sql += """
left outer join (
select
 Pn
,QtyS
,Noki
,Tanto
,Biko1
,KanDt
,KanQty
,NaraDt
from AcOrder
//where convert(KanQty,sql_decimal) = 0
where rtrim(KanQty) = '' or rtrim(KanQty) = '0'
) a on (a.Pn = o.HIN_GAI)
"""
        sql += """
where o.CANCEL_F <> '1'
and (
   (o.HAKKO_DT >= '{0}' and o.KAN_DT = '' and o.SHIJI_F <> '0')
""".format(self.HAKKO_DT)
        if self.UKEHARAI_CODE == 'ZN8' \
        or self.UKEHARAI_CODE == 'ZG7' \
        or self.UKEHARAI_CODE == 'ZH0' \
        or self.UKEHARAI_CODE == 'ZF0':
            sql += """
 or (o.HAKKO_DT >= '{0}' and o.KAN_DT = '' and o.SHIJI_F = '0')
""".format(self.HAKKO_DT)

        else:
#            sql += """
# or (o.HAKKO_DT >= '{0}' and o.KAN_DT = '' and o.SHIJI_F = '0' and (zqty92 > 0 or sumiQty92 > 0 or miQty92 > 0))
#""".format(self.HAKKO_DT)
            sql += """
 or (o.HAKKO_DT >= '{0}' and o.KAN_DT = '' and o.SHIJI_F = '0')
""".format(self.HAKKO_DT)
        if self.JGYOBU != "A":
            sql += """or (o.KAN_DT = '{0}' and (sumiQty92 > 0 or miQty92 > 0))""".format(self.KAN_DT)
        else:
            sql += """or (o.KAN_DT = '{0}')""".format(self.KAN_DT)
        sql += """)"""

        if self.JGYOBU != "":
            sql += " and o.JGYOBU='{0}'".format(self.JGYOBU)
        if self.TORI_KBN != "":
            sql += " and o.TORI_KBN='{0}'".format(self.TORI_KBN)
        if self.UKEHARAI_CODE in {'ZB0', 'PLAN'}:
            sql += " and o.SHIMUKE_CODE in ('01','02','03')"

        if sql_acnoki != "":
            sql += """
order by
 SHIJI_F desc
,acNoki
,o.SHIJI_NO
"""
        else:
            sql += """
order by
 zOrder desc
,SHIJI_F desc
,o.SHIJI_NO
"""
            
        return sql
#############
#メイン
#############
if __name__ == "__main__":
    p = p_sshiji()
    if 'REQUEST_METHOD' in os.environ:
        p.dns = get_req('dns','newsdc')
        p.JGYOBU = get_req('JGYOBU','')
        p.UKEHARAI_CODE = get_req('UKEHARAI_CODE','')
        p.top = int(get_req('top', 0))
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--JGYOBU", help="", default="", type=str)
        parser.add_argument("--UKEHARAI", help="", default="", type=str)
        parser.add_argument("--top", help="default: 0", default=0, type=int)
        parser.add_argument("--debug", action="store_true")
        args= parser.parse_args()
        _debug = args.debug
        p.dns= args.dns
        p.JGYOBU= args.JGYOBU
        p.UKEHARAI_CODE= args.UKEHARAI
        p.top= args.top
    p.main()
