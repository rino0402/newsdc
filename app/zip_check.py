# -*- coding: utf-8 -*-
import os
import sys
import io
import cgi
import cgitb
import re
import pandas as pd
import pyodbc
import json
import csv
import codecs
from datetime import date, datetime, timedelta
from decimal import Decimal
import traceback
from jusho import Jusho

def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)

def main(r):
    print("pyodbc.connect({0})".format(r["dsn"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dsn"])
    print("ok")

    num = _list(conn, r)

    print("conn.close()", end=".")
    conn.close()
    print("ok")
    
    return num

def _list(conn, r):
    sql = """
select
 SYUKA_YMD
,left(ID_NO,7) ID_NO_7
,YUBIN_No
,TEL_No
,OKURISAKI_CD
,OKURISAKI
,JYUSHO
,BIKOU
,min(DEN_NO) DEN_NO
,INS_BIN
,count(*) c
,UNSOU_KAISHA
,TYAKUTEN
,TYAKUTEN_NM
from {table}
where UNSOU_KAISHA='福山通運'
and SYUKA_YMD like '{dt}'
and YUBIN_No like '{_zip}'
group by
 SYUKA_YMD
,YUBIN_No
,TEL_No
,OKURISAKI_CD
,OKURISAKI
,ID_NO_7
,JYUSHO
,BIKOU
,INS_BIN
,UNSOU_KAISHA
,TYAKUTEN
,TYAKUTEN_NM
order by 1,2,3,4,5
""".format(table=r["table"],
           dt=r["dt"] if r["dt"] else "{:%Y%m%d}".format(datetime.now()),
           _zip=r["zip"]
           )
    df = pd.read_sql(sql, conn)
    print(df)

    num = 0
    for i, row in df.iterrows():
        for index, value in enumerate(row):
            if isinstance(value, str):
                row[index] = value.rstrip()
        print(row.YUBIN_No, end=" ")
        div_addr = divide_addess(row.JYUSHO.replace('　',''))
        print(div_addr[1], div_addr[2], div_addr[3])
        addr = address(row.JYUSHO)
        if addr:
            print(addr.postal_code, addr.prefecture_kanji, addr.city_kanji, addr.town_area_kanji)
            if addr.postal_code != row.YUBIN_No:
                eprint("★郵便番号エラー★")
                eprint(row.SYUKA_YMD, row.UNSOU_KAISHA,  row.TYAKUTEN, row.TYAKUTEN_NM)
                eprint(row.ID_NO_7, row.OKURISAKI_CD, row.OKURISAKI)
                eprint("★", row.YUBIN_No if row.YUBIN_No else "-------", div_addr[1], div_addr[2], div_addr[3])
                eprint("→", addr.postal_code, addr.prefecture_kanji, addr.city_kanji, addr.town_area_kanji)
                eprint("")
                num += 1
        else:
            print("■不明■")
            postman = Jusho()
            div_addr = divide_addess(row.JYUSHO.replace('　',''))
            for town in postman.towns_from_city(div_addr[1].strip(), div_addr[2].strip(), 'kanji'):
                print(town.town_area_kanji)
        print()
        #print(addr[1], addr[2], addr[3])
        #print(postman.address_from_town(addr[1], addr[2], addr[3], 'kanji'))
        #print(postman.address_from_town('東京都', '新宿区', '歌舞伎町', 'kanji'))
    return num

def address(address):
    postman = Jusho()
    addr = divide_addess(address.replace('　',''))
    #addr[2] = addr[2].strip()
    #print("{},{},{};".format(addr[1].strip(), addr[2].strip() , addr[3].strip()))
    ret_town = None
    while(ret_town == None):
        city = addr[2].strip()
        town_area = addr[3].strip()
        print(addr[1].strip(), city, 'kanji')
        for town in postman.towns_from_city(addr[1].strip(), city, 'kanji'):
            #print(town_area, town.town_area_kanji)
            #print(town.town_area_kanji)
            new_town = None
            if town_area.startswith(town.town_area_kanji) \
            or town_area.replace('大字', '').startswith(town.town_area_kanji) \
            or town_area.replace('字', '').startswith(town.town_area_kanji) \
            or town_area.replace('ヶ', 'ケ').startswith(town.town_area_kanji) \
            or town_area.replace('１', '一').startswith(town.town_area_kanji) \
            or town_area.replace('２条', '二条').startswith(town.town_area_kanji) \
            or town_area.replace('中川原中川原', '中川原').startswith(town.town_area_kanji) \
            or town_area.startswith(town.town_area_kanji.split('（')[0]) :
                #print(town.town_area_kanji, len(town.town_area_kanji))
                new_town = town
            """
            if new_town == None:
                if town.town_area_kanji in town_area \
                or town.town_area_kanji in town_area.replace('大字','') \
                or town.town_area_kanji in town_area.replace('ヶ','ケ') \
                or town.town_area_kanji in town_area.replace('２条','二条') :
                    new_town = town
            """
            if new_town:
                if len(new_town.town_area_kanji) > (len(ret_town.town_area_kanji) if ret_town else 0):
                    ret_town = new_town
            #print(new_town)
        if city.endswith('区') and len(addr.groups()) > 3:
            city += town_area
            town_area = addr[4].strip()
        else:
            break

    return ret_town
    #p = postman.address_from_town(addr[1].strip(), addr[2].strip(), addr[3].strip(), 'kanji')
    #return p

def divide_addess(address):
    pat = '(...??[都道府県])'
    pat += '((?:旭川|伊達|石狩|盛岡|奥州|田村|南相馬|那須塩原|東村山|武蔵村山|羽村|十日町|上越|富山|野々市|大町|蒲郡|四日市|姫路|大和郡山|廿日市|下松|岩国|田川|大村)市|.+?郡(?:玉村|大町|.+?)[町村]|.+?市.+?区|.+?[市区町村])'
    pat += '(.+)'

    matches = re.match(pat, address)
    #print(matches[1])
    #print(matches[2])
    #print(matches[3])
    return matches if matches else address
        
if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        pass
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dsn", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--dt", help="出荷日:yyyymmdd", default="", type=str)
        parser.add_argument("--table", help="y_syuka_h | del_syuka_h", default="y_syuka_h", type=str)
        parser.add_argument("--chaku", help="default:着店空白, *:全て", default="", type=str)
        parser.add_argument("--zip", help="郵便番号", default="%", type=str)
        num = main(vars(parser.parse_args()))
        print("num=", num)
        sys.exit(num)
