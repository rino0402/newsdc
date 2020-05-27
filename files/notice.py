# -*- coding: utf-8 -*-

import os
import json
import cgi
import datetime
import random

class notice:
    def __init__(self):
#       print('__init__')
        form = cgi.FieldStorage()
#        self.path = get_req('path','notice2')
        self.path = form.getvalue('path', 'notice2')
        self.status = 'success'
        self.filename = ''
        self.title = ''
        self.text = ''
        self.status = ''
        self.test = "test" in form

    def __del__(self):
#       print('__del__')
        pass
    
    def __str__(self):
        return self.path

    def getdir(self):
        abspath = os.path.dirname(os.path.abspath(__file__))
        abspath += '\\' + self.path
        directory = os.listdir(abspath)
#        print(directory)

    def get1file(self):
        path = os.path.dirname(os.path.abspath(__file__))
        path += '\\' + self.path
        try:
            for f in os.listdir(path):
                if os.path.isfile(path + '\\' + f):
                    self.filename = f
                    filepath = path + '\\' + f
                    file = open(filepath)
                    self.text = file.read()
                    file.close()
                    self.st_mtime = datetime.datetime.fromtimestamp(os.stat(filepath).st_mtime)
                    if self.test == False:
                        os.remove(path + '\\' + f)
        except:
            self.status = 'error'

    def print_response(self):
        r = {}
        r["status"] = self.status
        r["path"] = self.path
        r["filename"] = self.filename
        r["title"] = self.filename
        r["speech"] = ""
        r["chime"] = "chime"
        r["chime"] = "kamata4"
        r["chime"] = "tin1"
        r["volume"] = 0.1
        if self.filename == 'getoutn.ok':
            r["title"] = self.st_mtime.strftime('%H:%M') + " Active出荷データを受信しました."
            r["speech"] = self.st_mtime.strftime('%H:%M') + " Active,出荷データを受信しました."
            chime = ["kiikatsuura"]
            chime += ["haruyo1"]
            chime += ["atomu"]
            chime += ["kamata4"]
            chime += ["minatomirai1"]
            chime += ["ushiku1"]
            chime += ["ushiku2"]
            chime += ["999a"]
            chime += ["999b"]
            chime += ["999c"]
            chime += ["jinglebell-jingle"]
            r["chime"] = random.choice(chime)
        elif self.filename == 'getinn.txt':
            r["title"] = self.st_mtime.strftime('%H:%M') + " Active振替データを受信しました."
            r["speech"] = ""
            r["chime"] = "53 Dragon Quest 3 - Echoing Flute"
        elif self.filename == 'getoutg.txt':
            r["title"] = self.st_mtime.strftime('%H:%M') + " Glics出荷データを受信しました."
            r["speech"] = ""
            r["chime"] = "50 Dragon Quest 3 - Fanfare 1"
        elif self.filename == 'geting.txt':
            r["title"] = self.st_mtime.strftime('%H:%M') + " Glics振替データを受信しました."
            r["speech"] = ""
            r["chime"] = "55 Dragon Quest 3 - Silver Harp"
        elif self.filename == 'pos_start.txt':
            r["title"] = self.st_mtime.strftime('%H:%M') + " POSシステムを起動しました."
            r["speech"] = self.st_mtime.strftime('%H:%M') + " ポスシステムを起動しました."
            r["chime"] = "45 Dragon Quest 3 - Wayfarer's Inn"
        elif self.filename == 'y_syuka_check.txt':
            r["title"] = self.st_mtime.strftime('%H:%M') + " Active出荷実績データを送信しました."
#            r["speech"] = self.st_mtime.strftime('%H:%M') + " Active,出荷実績データを送信しました."
            r["speech"] = "。"
            r["chime"] = "48 Dragon Quest 3 - Special Item"
        elif self.filename != '':
            r["title"] = self.st_mtime.strftime('%H:%M ') + self.filename
        r["text"] = self.text
        r["test"] = self.test
        print('Content-Type:application/json; charset=UTF-8;\n')
        print(json.dumps(r, ensure_ascii=True, indent=4, sort_keys=False, separators=(',', ': ')))

if __name__ == "__main__":
    n = notice()
#    print(n)
#   n.getdir()
    n.get1file()
    n.print_response()
#    path = os.path.dirname(os.path.abspath(__file__)) 
#    path += '\\notice2'
#    print(path)
#    directory = os.listdir(path)
#    print(directory)
