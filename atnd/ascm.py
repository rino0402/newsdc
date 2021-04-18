# -*- coding: utf-8 -*-
import os
import sys
import io
import csv
import pandas as pd
import pyodbc

def main(r):
    print("main({})".format(r))
    print("pyodbc.connect({0})".format(r["dns"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dns"])
    print("ok")

    with open(r["csv"], "r", encoding="Shift-JIS", errors='ignore') as file:
        df = pd.read_table(file, delimiter=",")
        print(df)
        df = df.fillna("")
        print("{} {}".format(df["日付"].min().replace("/","-"), df["日付"].max().replace("/","-")))
        sql = """delete from Ascm
where Dt between '{}' and '{}'
""".format(df["日付"].min().replace("/","-"), df["日付"].max().replace("/","-"))
        print(sql)
        conn.execute(sql)
        for i, row in df.iterrows():
            for c, col in enumerate(row):
                if isinstance(col, str):
                    row[c] = col.rstrip()
            print(row)
            insert(conn, row)

    conn.commit()
    print("conn.close()", end=".")
    conn.close()
    print("ok")

def insert(conn, row):
    sql = """insert into Ascm (
	StaffNo	
,	Name	
,	Dt		
,	Kubun	
,	Awh		
,	Shift	
,	ShiftNm	
,	BegTm	
,	FinTM	
,	Late	
,	Early	
,	Extra	
,	Night	
,	H1Extra	
,	H1Night	
,	H2Extra	
,	H2Night	
,	PTO		
,	Actual	
,	Memo
) values (
"""
    sql += " '{:0>5}'".format(row["社員No"])
    sql += ",'{}'".format(row["氏名"])
    sql += ",'{}'".format(row["日付"]).replace("/","-")
    sql += ",'{}'".format(row["区分"])
    sql += ",'{}'".format(float(row["実働"].replace(":",".") or 0))
    sql += ",'{}'".format(str(row["シフトNo"])[0] if row["シフトNo"] else "")
    sql += ",'{}'".format(str(row["ｼﾌﾄ名"])[0] if row["ｼﾌﾄ名"] else "")
    sql += ",'{}'".format(row["出勤"]) if row["出勤"] else ",null"
    sql += ",'{}'".format(row["退勤"]) if row["退勤"] else ",null"
    sql += ",'{}'".format(float(row["遅刻"].replace(":",".") or 0))
    sql += ",'{}'".format(float(row["早退"].replace(":",".") or 0))
    sql += ",'{}'".format(float(row["普通残業"].replace(":",".") or 0))
    sql += ",'{}'".format(float(row["深夜残業"].replace(":",".") or 0))
    sql += ",'{}'".format(float(row["法定休残"].replace(":",".") or 0))
    sql += ",'{}'".format(float(row["法定休深"].replace(":",".") or 0))
    sql += ",'{}'".format(float(row["所定休残"].replace(":",".") or 0))
    sql += ",'{}'".format(float(row["所定休深"].replace(":",".") or 0))
    sql += ",'{}'".format(row["有給休暇"])
    sql += ",'{}'".format(float(row["普通時間"].replace(":",".") or 0))
    sql += ",'{}'".format(row["備考"])
    sql += ")"
    print(sql)
    conn.execute(sql)
    return conn.execute("select @@rowcount").fetchone()[0]
    
if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        pass
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("csv", help="", default="", nargs='?')
        r = main(vars(parser.parse_args()))
