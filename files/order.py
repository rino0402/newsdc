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

class order:
    dns = "newsdc"
    top = 0
    sql = ""
    data = {}
    results = {}
    _table = "Y_Syuka"

    @property
    def table(self):
        return str(self._table)
    @table.setter
    def table(self, value):
        if value != "":
            self._table = "(select * from del_syuka where KEY_SYUKA_YMD like '{0}')".format(value)

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
        self.make_list()
        debug("close()")
        self.conn.close()

    def get_sql(self):
        debug(self.__class__.__name__ + '.' + sys._getframe().f_code.co_name)
        return """
select
 y.KEY_SYUKA_YMD
,y.KEY_MUKE_CODE
,y.LK_MUKE_CODE
,y.MUKE_NAME
,y.KEY_CYU_KBN
,y.CYU_KBN
,y.CYU_KBN_NAME
,y.KEY_ID_NO
,y.ODER_NO
,y.JGYOBU
,y.NAIGAI
,y.KEY_HIN_NO
,y.HIN_NAME yHIN_NAME
,i.HIN_NAME iHIN_NAME
,convert(y.SURYO,sql_decimal) qty
,y.KAN_KBN
,y.KENPIN_YMD
,y.KAN_YMD
,y.KENPIN_YMD
,i.L_KISHU3
,case
 when i.L_KISHU3 like 'MADE%' then i.L_KISHU3
 when i.INSP_MESSAGE  like '%リチウム%' then i.INSP_MESSAGE
 when y.LK_MUKE_CODE like '103%' then '韓国リサイクルマーク個別表示'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('AMV79V-L7','AMC39V-EDT') then '台湾 RoHS+標示法'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('AMV79V-L70K','AMC39VEDT00J','AMV79V-L70K','AMC39VEDT00J') then '台湾 RoHS'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('AMV88M-L70K') then '台湾 標示法(AMV44M-L7)'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('ARK06-E1800S','ASR792-454DK') then '台湾 銘板(KRY303E18)'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('ARK02-55900S','ARK02-5592') then '台湾 銘板(KRY303559)'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('ARK02-19700S','ASR79W-281AK') then '台湾 銘板(KRY303559)'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('ARK16-A0600S','ARK16EA06','ARK16E722') then '台湾 銘板(KRY303A06)'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('ARC00-D1500S','ARC00-D15J2U') then '台湾 銘板(KRY303D15)'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('ARC00-F8200S','ARC00-F82HBU') then '台湾 銘板(KRY303855)'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('ARC00-85500S','ARC00-855PAU') then '台湾 銘板(KRY303855)'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('ARK02-B47K3S') then '台湾 銘板(KRY303559)'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('ARC00-F44KHS') then '台湾 銘板(KRY303855)'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('APN07B628-0U') then '番號：PN07B628100U'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('APN07B48810U') then '番號：PN07B488300U'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('AZN09BD88-KS') then '台湾 銘板(ZY56BD88)'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('ARQ53TA0600S') then '台湾 銘板(RY48TA06)'
 when y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in ('ARQ53TE9300S','ARQ533G4000S','ARQ53TJ8500S') then '台湾 銘板(RY48TE93)'
 when y.LK_MUKE_CODE like '106%' then '台湾 標示法'
 end Careful
from {0} y
left outer join item i on (y.JGYOBU = i.JGYOBU and y.NAIGAI = i.NAIGAI and y.KEY_HIN_NO = i.HIN_GAI)
where KEY_CYU_KBN = 'E'
and (
   (i.L_KISHU3 like 'MADE%')
or (i.INSP_MESSAGE like '%リチウム%')
or (y.LK_MUKE_CODE like '103%')
or (y.LK_MUKE_CODE like '106%' and y.KEY_HIN_NO in (
 'AMV87TLS0V0J'
,'AMV91TLS0V0J'
,'AMV84K-JS0H'
,'AMV97V-L7'
,'AMV79V-L7'
,'AMC79VECR0V'
,'AMC39V-EDT'
,'AMC85KEDU0RJ'
,'AMC95BEDU0RJ'
,'AMC37KECR0V'
,'AMC38KECR0V'
,'AMC95KECR0V'
,'AMV92L-6S03'
,'AMV77J-CQ0'
,'AMV30K-AV0'
,'AMV86LDH000J'
,'AMV0VKDH000J'
,'AMV95K-AT0'
,'AMV44M-L7'
,'AMV79V-L7','AMC39V-EDT'
,'AMV79V-L70K','AMC39VEDT00J','AMV79V-L70K','AMC39VEDT00J'
,'AMV88M-L70K'
,'ARK06-E1800S','ASR792-454DK'
,'ARK02-55900S','ARK02-5592'
,'ARK02-19700S','ASR79W-281AK'
,'ARK16-A0600S','ARK16EA06','ARK16E722'
,'ARC00-D1500S','ARC00-D15J2U'
,'ARC00-F8200S','ARC00-F82HBU'
,'ARC00-85500S','ARC00-855PAU'
,'ARK02-B47K3S'
,'ARC00-F44KHS'
,'AVV43K-QQ0','AVV97V-QYT','AVV92K-QQ0H','AVV0VK-QQ0S','AVV61V-QYT'
,'AVV88C-NF0K','AVV93TRA0K0J','AVV84R-RA0B','AVV97R-RA0K','AVV0YA-RA0W','AVV00K-RA0V','AVV36P-NF0H','AVV88K-RA0V','AVV38K-RA02'
,'AVV92K-RA0V','AVV95K-RA0V','AVV97V-RAT','AVV27A-RA02','AVV79V-RAT','AVV61V-RAT'
,'APN07B628-0U','APN07B48810U','AZN09BD88-KS','ARQ53TA0600S','ARQ53TE9300S','ARQ533G4000S','ARQ53TJ8500S'
)
)
)
order by
 KEY_SYUKA_YMD
,KEY_MUKE_CODE
,LK_MUKE_CODE
,CYU_KBN
,CYU_KBN_NAME
,ODER_NO
""".format(self.table)

    def make_list(self):
        debug(self.__class__.__name__ + '.' + sys._getframe().f_code.co_name)
        data = []
        sql= self.get_sql()
        cur = self.conn.cursor()
        debug(sql)
        cur.execute(sql)
        columns = [column[0] for column in cur.description]
        for c in cur.fetchall():
            for i, v in enumerate(c):
                if isinstance(v, str):
                    c[i]=v.rstrip()
            dic = dict(zip(columns, c))
            debug(dic)
            data.append(dic)
        results = {}
        results["data"] = data
        print('Content-Type:application/json; charset=UTF-8;\n')
        print(json.dumps(results,  \
                             ensure_ascii=True, indent=4, sort_keys=False, \
                             separators=(',', ': ') \
                             , cls=DatetimeEncoder \
                             ) \
                  )

    def select_top(self):
        debug(self.__class__.__name__ + '.' + sys._getframe().f_code.co_name)
        if self.top > 0:
            return "top {0}".format(self.top)
        return ""

    def print_response(self):
        results = {}
        results["dns"] = self.dns
        results["sql"] = self.sql
        results["data"] = self.data
        print('Content-Type:application/json; charset=UTF-8;\n')
        #default=decimal_default_proc,
        print(json.dumps(results,  \
                         ensure_ascii=True, indent=4, sort_keys=False, \
                         separators=(',', ': ') \
                         , cls=DatetimeEncoder \
                         ) \
              )

def decimal_default_proc(obj):
    if isinstance(obj, Decimal):
        return float(obj)
    raise TypeError

class DatetimeEncoder(json.JSONEncoder):
    def default(self, obj):
        debug(type(obj))
        if isinstance(obj, str):
            return 'a'
        if isinstance(obj, datetime ):
            return obj.strftime("%Y-%m-%d %H:%M:%S")
        if isinstance(obj, date ):
            return obj.strftime("%Y-%m-%d")
        if isinstance(obj, Decimal):
            return float(obj)
        return json.JSONEncoder.default(self, obj)

def get_req(nm, v):
    form = cgi.FieldStorage()
    if nm in form:
        v = form[nm].value
    return v

#############
#メイン
#############
if __name__ == "__main__":
    p = order()
    if 'REQUEST_METHOD' in os.environ:
        p.dns = get_req('dns','newsdc')
        p.table = get_req('table','')
        p.top = int(get_req('top', 0))
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--top", help="default: 0", default=0, type=int)
        parser.add_argument("--table", help="ex.20180620", default=0, type=str)
        parser.add_argument("--debug", action="store_true")
        args= parser.parse_args()
        _debug = args.debug
        p.dns= args.dns
        p.top= args.top
        p.table= args.table
    p.main()
