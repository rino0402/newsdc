# -*- coding: utf8 -*-
import os
import sys
import re
import tkinter
from tkinter import ttk
import usb.core
from usb.backend import libusb1
import bluetooth
import configparser

#ラベル発行
def lavel(text):
    print("lvael({})".format(text))
#    print("libusb1.get_backend()")
#    be = libusb1.get_backend(find_library=lambda x: "J:\newsdc\files\dist\libusb-1.0.dll")
#    be = libusb1.get_backend()
#    print(be)
#    print("usb.core.find()")
#    dev =  usb.core.find(backend=be, idVendor=0xfa11, idProduct=0xfe25)
    dev =  usb.core.find(idVendor=0xFA11, idProduct=0xFE25)
#    print(dev)
    """
    try:
        cfg = dev.get_active_configuration()
        print(cfg)
    except usb.core.USBError:
        cfg = None
    if cfg is None: #or cfg.bConfigurationValue != cfg_desired:
        dev.set_configuration()
    """
    dev.set_configuration()
    ep = 1
    dev.write(ep, text)
    dev.reset()

#ラベル発行:Bluetooth
def print_bluetooth(text):
    print("print_bluetooth()")
    prgBar.start()
    prgTxt["text"] = "BT:印刷中..."
    root.update()
    addr = textAddr.get()
    if addr == "":
        addr = None
    print("bluetooth.find_service:{}.".format(addr), end="")
    service_matches = bluetooth.find_service(address=addr, name=b'Serial Port DevB')
    print(len(service_matches))
    if len(service_matches) == 0:
        prgTxt["text"] = "BT:接続エラー"
        prgBar.stop()
        return
    for sv in service_matches:
        print(sv)
       
    first_match = service_matches[0]
    port = first_match["port"]
    name = first_match["name"]
    host = first_match["host"]
    print("{} {}".format(name, host))
    
    sock = bluetooth.BluetoothSocket(bluetooth.RFCOMM)
    sock.connect((host, port))
    sock.send(text)
    sock.close()
    if addr is None:
        textAddr.insert(tkinter.END, host)
    prgTxt["text"] = "BT:印刷完了"
    prgBar.stop()
#ボタン
def OkButton(event):
    print(event)
#    labelFoot["text"] = "印刷中..."
    prgBar.start()
    prgTxt["text"] = "印刷中..."
    dispImage(event)
    nstxt = ""
    nsbcd = ""
#    text = "{: ^15}".format(dispText["text"]).rstrip()
    text = dispText["text"].rstrip()
    x = max(0, int((344 - len(text) * 46/2) / 2))
    bcd = "{}{}".format(textTop.get(), text1.get().rstrip())
    if len(bcd) < 10:
        nw = 2
        bx = max(0, int((344 - len(bcd) * 40) / 2))
        mg = 1
    elif len(bcd) < 14:
        nw = 1
        bx = max(0, int((344 - len(bcd) * 20) / 2))
        mg = 1
    else:
        nw = 1
        bx = max(0, int((344 - len(bcd) * 20) / 2))
        mg = 0
            
    ns = textNS.get()
    if int(ns or 0) > 0:
        nstxt = ",NS={},NE={}".format(len(text) - int(ns) + 1,ns)
        nsbcd = ",NS={},NE={}".format(len(bcd) - int(ns) + 1,ns)
    txt = '''
JOB
DEF MK=1,MD=1,DR=2,DK={dk},MS=0,GS=0,PO=0,TO=0,PW=384,PH=344,PG=16,UM=24,BM=0,XO=0,AF=1,AB=0
START
#FONT TP=7,DR=1,CS=0,LS=0,WD=45,LG=120,SL=0
FONT TP=7,WD=46,LG=120
TEXT X={x},Y=45,L=1{nstxt}
{text}
#CD TP=7,X=0,Y=165,DR=1,NW=2,RA=2,HT=100,CD=0,HR=1,MG=1,BX=0,SC=0,EC=0
BCD TP=7,X={bx},Y=165,DR=1,NW={nw},RA=2,HT=100,CD=0,HR=1,MG={mg}{nsbcd}
{bcd}
QTY P={p}
END
DEF MK=1,MD=3,PH=344,PW=384,UM=24,BM=0,DK=12,XO=8,AF=1,MS=28,PO=25,TO=100,PG=24
JOBE
'''.format(text=text,
           x=x,
           bcd=bcd,
           bx=bx,
           mg=mg,
           nw=nw,
           nstxt=nstxt,
           nsbcd=nsbcd,
           p=textQty.get(),
           dk=textDK.get())
    print(txt)
    if varPrinter.get() == 0:
        lavel(txt)
    else:
        print_bluetooth(txt)
#    labelFoot["text"] = "印刷完了."
#    prgBar.stop()
    text1.focus_set()
    
#ラベル表示イメージ
def dispImage(event):
    dispText["text"] = text1.get()
    if textSep.get() != "":
        dispText["text"] = textSep.get().join(re.split('(..)',dispText["text"])[1::2])
    barCode["text"] = "*{}{}*".format(textTop.get(), text1.get())
    barText["text"] = barCode["text"] 
    root.update()
    
#Text
def text1FocusOut(event):
    print(event)
    print(text1.get())
    dispImage(event)

def get_config():
    return os.path.dirname(os.path.abspath(__file__)) + '\\Davinci.cfg'

cfg = configparser.ConfigParser()
cfg.read(get_config())

root = tkinter.Tk()
root.title("DaVinciラベル発行 - 0.02")
# 1
frame1 = tkinter.Frame(padx=5, pady=5)
frame1.pack()
#ラベル
label1 = tkinter.Label(frame1, text='棚番')
label1.pack(side="left")
#入力
text1 = tkinter.Entry(frame1, font=("", 24), width=16)
try:
    text1.insert(tkinter.END, cfg['Davinci']['text'])
except KeyError:
    text1.insert(tkinter.END, "1A010101")
text1.bind("<FocusOut>", text1FocusOut)
text1.pack()
# 2
frame2 = tkinter.Frame(padx=5, pady=5)
frame2.pack()
label2 = tkinter.Label(frame2, text='先頭')
label2.pack(side="left")
textTop = tkinter.Entry(frame2, font=("", 24), width=1)
try:
    textTop.insert(tkinter.END, cfg['Davinci']['top'])
except KeyError:
    textTop.insert(tkinter.END, "/")
textTop.bind("<FocusOut>", dispImage)
textTop.pack(side="left")
#textTop.grid(column=1, row=1)
label3 = tkinter.Label(frame2, text='区切り')
label3.pack(side="left")
textSep = tkinter.Entry(frame2, font=("", 24), width=1)
try:
    textSep.insert(tkinter.END, cfg['Davinci']['sep'])
except KeyError:
    textSep.insert(tkinter.END, "-")
textSep.bind("<FocusOut>", dispImage)
textSep.pack(side="left", fill="x", expand=True)

labelNS = tkinter.Label(frame2, text='連番')
labelNS.pack(side="left")
textNS = tkinter.Entry(frame2, font=("", 24), width=1)
try:
    textNS.insert(tkinter.END, cfg['Davinci']['ns'])
except KeyError:
    textNS.insert(tkinter.END, "2")
textNS.bind("<FocusOut>", dispImage)
textNS.pack(side="left", fill="x", expand=True)
labelNS2 = tkinter.Label(frame2, text='2')
labelNS2.pack(side="left")

#textSep.grid(column=2, row=1)
# 3
frame3 = tkinter.Frame(padx=5, pady=5, borderwidth=3, relief="solid")
frame3.pack()
dispText = tkinter.Label(frame3, font=("", 24))
dispText.pack()
#dispText.grid(column=1, row=2)
barCode = tkinter.Label(frame3, text="*BARCODE*", font=("3 of 9 Barcode", 24))
barCode.pack()
#barCode.grid(column=1, row=3)
barText = tkinter.Label(frame3, text="*BARCODE*", font=("", 12))
barText.pack()
#barText.grid(column=1, row=4)

# 4
frame4 = tkinter.Frame(padx=5, pady=5)
frame4.pack()
label4 = tkinter.Label(frame4, text='枚数')
label4.pack(side="left")
textQty = tkinter.Entry(frame4, font=("", 24), justify="center", width=3)
try:
    textQty.insert(tkinter.END, cfg['Davinci']['qty'])
except KeyError:
    textQty.insert(tkinter.END, "1")
textQty.pack(side="left")

# 5 発行ボタン
button1 = tkinter.Button(text='発行', font=("", 24), relief="solid",
                         padx=5, pady=5,bg="yellow")
button1.bind("<Button-1>", OkButton)
button1.bind("<Return>", OkButton)
button1.pack()
prgBar = ttk.Progressbar()
prgBar.pack()
prgTxt = tkinter.Label(text='')
prgTxt.pack()
#button1.grid(column=1, row=5)
# option
frameOpt = tkinter.Frame(padx=5, pady=5)
frameOpt.pack()
labelDK = tkinter.Label(frameOpt, text='印字濃度')
labelDK.pack(side="left")
textDK = tkinter.Entry(frameOpt, font=("", 14), justify="right", width=3)
try:
    dk = (cfg['Davinci']['dk'] or "12")
except KeyError:
    dk = "12"
textDK.insert(tkinter.END, dk)
textDK.pack(side="left")
labelDK2 = tkinter.Label(frameOpt, text='1-16')
labelDK2.pack(side="left")
# printer
framePrinter = tkinter.Frame(padx=5, pady=5)
framePrinter.pack()
varPrinter = tkinter.IntVar()
try:
    varPrinter.set(cfg['Davinci']['printer'])
except KeyError:
    varPrinter.set(0)
rdoUsb = tkinter.Radiobutton(framePrinter, value=0, variable=varPrinter, text='USB',
                             anchor='w', font=("", 12))
rdoUsb.pack(side="left")
rdoBt = tkinter.Radiobutton(framePrinter, value=1, variable=varPrinter, text='BT',
                            anchor='w', font=("", 12))
rdoBt.pack(side="left")
textAddr = tkinter.Entry(framePrinter, font=("", 12), justify="left", width=len("00:01:90:EE:87:18"))
try:
    textAddr.insert(tkinter.END, cfg['Davinci']['addr'])
except KeyError:
    textAddr.insert(tkinter.END, "")
textAddr.pack(side="left")

# footer
frameFoot = tkinter.Frame(padx=5, pady=5)
frameFoot.pack()
labelFoot = tkinter.Label(frameFoot, text='0.02 Bluetooth対応',
                          width=40, anchor='w', justify='left')
labelFoot.pack()

dispImage(None)
def on_closing():
    cfg['Davinci'] = {
        'text': text1.get(),
        'top': textTop.get(),
        'sep': textSep.get(),
        'ns': textNS.get(),
        'qty': textQty.get(),
        'dk': textDK.get(),
        'printer': varPrinter.get(),
        'addr': textAddr.get()
    }
    with open(get_config(), "w") as config_file:
        cfg.write(config_file)
    root.destroy()
    
root.protocol("WM_DELETE_WINDOW", on_closing)    
root.mainloop()
