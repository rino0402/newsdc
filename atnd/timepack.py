# -*- coding: utf-8 -*-
import os
import sys
import io
import csv
import pyodbc
from datetime import date, datetime, timedelta, time
from dateutil.relativedelta import relativedelta
#-------------------------------------------------------------
def main(r):
    print("main({})".format(r))
    print("pyodbc.connect({})".format(r["dns"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dns"])
    print("ok")
    sql_del = ""
    with open(r["csv"], "r") as f:
        #reader = csv.reader(f)
        part = False
        for row in csv.reader(f):
            print(len(row), row)
            if row[0] == "カード番号":
                if row[29] == "コメント":
                    part = True
                else:
                    part = False
                continue
            d = {}
            d["CardNo"] = "{}".format(row[0])
            d["StaffNo"] = "{}".format(row[1])
            d["Name"] = "{}".format(row[2])
            d["Post"] = "{}".format(row[3])
            d["Dt"] = "{}".format(row[4].replace("/","-"))
            if sql_del == "":
                dt1 = datetime.strptime(d["Dt"], '%Y-%m-%d')
                if dt1.day < 16:
                    dt1 -= relativedelta(months=1)
                dt1 = date(dt1.year,dt1.month,16)
                dt2 = dt1 + relativedelta(months=1)
                dt2 = date(dt2.year,dt2.month,15)
                sql_del = "delete from timepack where Dt between '{:%Y-%m-%d}' and '{:%Y-%m-%d}'".format(dt1, dt2)
                print(sql_del, end=".")
                conn.execute(sql_del)
                print(conn.execute("select @@rowcount").fetchone()[0])
                
            d["Shift"] = "{}".format(row[5])
            d["Holiday"] = "{}".format(row[6])
            d["Absence"] = "{}".format(row[7])
            if row[8]:
                d["BgnTM"] = "{}".format(row[8])
            d["BgnMK"] = "{}".format(row[9])
            if row[10]:
                d["OutTM"] = "{}".format(row[10])
            d["OutMK"] = "{}".format(row[11])
            if row[12]:
                d["BckTM"] = "{}".format(row[12])
            d["BckMK"] = "{}".format(row[13])
            if row[14]:
                d["FinTM"] = "{}".format(row[14])
            d["FinMK"] = "{}".format(row[15])
            if row[16]:
                d["Ex1TM"] = "{}".format(row[16])
            d["Ex1MK"] = "{}".format(row[17])
            if row[18]:
                d["Ex2TM"] = "{}".format(row[18])
            d["Ex2MK"] = "{}".format(row[19])
            d["Actual"] = "{}".format(row[20])
            d["Extra"] = "{}".format(row[21])
            d["ExtEarly"] = "{}".format(row[22])
            d["Night"] = "{}".format(row[23])
            if part == False:
                d["ExtNight"] = "{}".format(row[24])
                d["Holiday1"] = "{}".format(row[25])
                d["HolidayNight1"] = "{}".format(row[26])
                d["Holiday2"] = "{}".format(row[27])
                d["HolidayNight2"] = "{}".format(row[28])
                d["LateEarly"] = "{}".format(row[29])
                d["Private"] = "{}".format(row[30])
                d["Memo"] = "{}".format(row[31])
            else:
                d["Holiday1"] = "{}".format(row[24])
                d["Holiday1"] = "{}".format(row[24])
                d["HolidayNight1"] = "{}".format(row[25])
                d["Holiday2"] = "{}".format(row[26])
                d["HolidayNight2"] = "{}".format(row[27])
                d["Private"] = "{}".format(row[28])
                d["Memo"] = "{}".format(row[29])
            """
,	ExtEarly		Currency default 0	not null	//22"早出残業"	,22"深夜時間"
,	Night			Currency default 0	not null	//23"深夜時間"	,23"基準外深夜"
,	ExtNight		Currency default 0	not null	//24"深夜残業"	,24"休１時間"
,	Holiday1		Currency default 0	not null	//25"休１時間"	,25"休１深夜"
,	HolidayNight1	Currency default 0	not null	        //26"休１深夜"	,26"休２時間"
,	Holiday2		Currency default 0	not null	//27"休２時間"	,27"休２深夜"
,	HolidayNight2	Currency default 0	not null	        //28"休２深夜"	,28"外出時間"
,	LateEarly		Currency default 0	not null	//29"遅早時間"	,29"コメント","","","","","","",""
,	Private			Currency default 0	not null	//30"外出時間"
,	Memo			Char( 2) default '' not null	        //31"コメント","","","","",""

            """
            d["insert"] = insert(conn, d)
            print(d)
    print("conn.commit()", end=".")
    conn.commit()
    print("ok")

    print("conn.close()", end=".")
    conn.close()
    print("ok")
    return r
#-------------------------------------------------------------
def insert(conn, data):
    sql = "insert into TimePack (" + ",".join(map(str, data)) + ") values ("
    for d in data:
        sql += "'{0}',".format(data[d])
    sql = sql[:-1] + ")"
    print(sql, end=".")
    try:
        conn.execute(sql)
        ret = conn.execute("select @@rowcount").fetchone()[0]
        print(ret)
        return ret
    except pyodbc.ProgrammingError as e:
        print(e)
        raise
    except pyodbc.Error as e:
        if '(Btrieve Error 5) (-4994)' in str(e.args):
            print("(W)")
            print(e)
            raise
        else:
            print(e)
            raise
    except:
        print('error')
        raise
#-------------------------------------------------------------
if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        pass
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("csv", help="", default="", nargs='?')
        r = main(vars(parser.parse_args()))
