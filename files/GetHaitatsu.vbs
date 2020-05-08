<%
Const vbWide=4 	'文字列内の半角文字を全角文字に変換
Private Function GetVersion()
	GetVersion = "2016.06.07 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～千葉県)"
	GetVersion = "2016.06.08 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～神奈川県)"
	GetVersion = "2016.06.09 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～山梨県)"
	GetVersion = "2016.06.13 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～新潟県)"
	GetVersion = "2016.06.14 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～長野県)"
	GetVersion = "2016.06.15 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～富山県)"
	GetVersion = "2016.06.16 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～福井県)"
	GetVersion = "2016.06.17 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～静岡県)"
	GetVersion = "2016.06.20 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～岐阜県)"
	GetVersion = "2016.06.22 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～愛知県)"
	GetVersion = "2016.06.23 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～三重県、滋賀県)"
	GetVersion = "2016.06.24 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～京都)"
	GetVersion = "2016.06.27 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～和歌山)"
	GetVersion = "2016.06.28 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～奈良県)"
	GetVersion = "2016.06.29 出力形式：集計表(送り先別)：配達不可(対応中...済 北海道～山口)"
	GetVersion = "2016.06.30 出力形式：集計表(送り先別)：福通配達不可対応済 北海道～沖縄(九州除く)"
	GetVersion = "2018.03.14 出力形式：集計表(送り先別)：福通配達不可対応2018 山形県"
	GetVersion = "2018.12.03 出力形式：集計表(送り先別)：福通配達不可対応2019 静岡県、岐阜県"
	GetVersion = "2019.06.27 出力形式：集計表(送り先別)：福通配達不可対応2019.5.1 北海道"
	GetVersion = "2019.06.28 出力形式：集計表(送り先別)：福通配達不可対応2019.5.1 岩手県"
	GetVersion = "<span class=""new"">2019.08.26 出力形式：集計表(送り先別)：福通配達不可対応2019.8.1 埼玉県／福島県</span>"
	GetVersion = "<span class=""new"">2019.09.30 出力形式：集計表(送り先別)：福通配達不可対応2019.10.1 宮城県</span>"
	GetVersion = "<span class=""new"">2019.12.23 配達不能エリア(19.10.1) 四国(4県)、鳥取、島根</span>"
	GetVersion = "<span class=""new"">2019.12.27 配達不能エリア(20.1.1) 福島、宮崎、沖縄</span>"
	GetVersion = "<span class=""new"">2020.01.06 配達不能エリア(20.1.1) 熊本、大分</span>"
	GetVersion = "<span class=""new"">2020.01.08 配達不能エリア(20.1.1) 千葉 ※正規表現</span>"
	GetVersion = "<span class=""new"">2020.01.09 配達不能エリア(20.1.1) 山形、秋田</span>"
	GetVersion = "<span class=""new"">2020.01.09 配達不能エリア(20.1.1) 処理速度改善</span>"
	GetVersion = "<span class=""new"">2020.01.10 配達不能エリア(20.1.1) 山梨、新潟</span>"
	GetVersion = "<span class=""new"">2020.01.14 配達不能エリア(20.1.1) 長野、富山</span>"
	GetVersion = "<span class=""new"">2020.01.20 配達不能エリア(20.1.1) 静岡</span>"
	GetVersion = "<span class=""new"">2020.01.21 配達不能エリア(20.1.1) 岐阜</span>"
	GetVersion = "<span class=""new"">2020.01.22 配達不能エリア(20.1.1) 愛知</span>"
	GetVersion = "<span class=""new"">2020.01.23 配達不能エリア(20.1.1) 三重</span>"
	GetVersion = "<span class=""new"">2020.01.24 配達不能エリア(20.1.1) 滋賀,京都,和歌山,兵庫,奈良,岡山,広島</span>"
	GetVersion = "<span class=""new"">2020.01.28 配達不能エリア(20.1.1) 鳥取,島根</span>"
	GetVersion = "<span class=""new"">2020.01.30 配達不能エリア(20.1.1) 山口,徳島,香川,愛媛,高知</span>"
	GetVersion = "<span class=""new"">2020.04.07 配達不能エリア(2020.3.23) 東京</span>"
'	xcopy/d/y GetHaitatsu.vbs \\w5\newsdc\files\
End Function
Function GetHaitatsu(byVal strAddress)
	strAddress = Han2Zen(strAddress)
	strAddress = Replace(strAddress," ","")
	strAddress = Replace(strAddress,"　","")
	GetHaitatsu = ""
	'北海道
	GetHaitatsu = Hokaido(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'岩手県
	GetHaitatsu = Iwate(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'秋田県
	GetHaitatsu = Akita(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'宮城県 2019.09.30
	GetHaitatsu = Miyagi(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'山形県
	GetHaitatsu = Yamagata(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'福島県
	GetHaitatsu = Fukushima(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'茨城県
	if inStr(strAddress,"鹿島市") then
	if inStr(strAddress,"光３番") _
	or inStr(strAddress,"光２番") _
	or inStr(strAddress,"光４番") _
	or inStr(strAddress,"新浜５番") _
	or inStr(strAddress,"新浜２１番") _
								  then
		GetHaitatsu = "茨城県 鹿島市（光３番地新日鐵住金㈱、光２番地新日鉄住金ステンレス㈱、光４番地中央電気工業㈱、新浜５番地鹿島共同火力㈱、新浜２１番地新浜工業団地）、神栖市（光１番地日鉄住金物流㈱）<br>支店止め又はチャーター"
		exit function
	end if
	end if
	if inStr(strAddress,"神栖市") then
	if inStr(strAddress,"光１番") _
								  then
		GetHaitatsu = "茨城県 鹿島市（光３番地新日鐵住金㈱、光２番地新日鉄住金ステンレス㈱、光４番地中央電気工業㈱、新浜５番地鹿島共同火力㈱、新浜２１番地新浜工業団地）、神栖市（光１番地日鉄住金物流㈱）<br>支店止め又はチャーター"
		exit function
	end if
	end if
	if inStr(strAddress,"つくば市") then
	if inStr(strAddress,"筑波") _
	or inStr(strAddress,"沼田") _
	or inStr(strAddress,"国松") _
								  then
		GetHaitatsu = "茨城県 つくば市筑波、沼田、国松<br>支店止め又はチャーター"
		exit function
	end if
	end if
	if inStr(strAddress,"つくばみらい市") then
	if inStr(strAddress,"筒戸１４１０") _
								  then
'		GetHaitatsu = "茨城県 つくばみらい市筒戸1410 センコー㈱<br>支店止め又はチャーター"
		exit function
	end if
	end if
	if inStr(strAddress,"常総市") then
	if inStr(strAddress,"内守谷町") then
	if inStr(strAddress,"きぬの里") then
	if inStr(strAddress,"３") then
	if inStr(strAddress,"３８") then
	if inStr(strAddress,"１") then
		GetHaitatsu = "茨城県 常総市内守谷町きぬの里3-38-1 日本郵便㈱<br>支店止め又はチャーター"
		exit function
	end if
	end if
	end if
	end if
	end if
	end if
'	if inStr(strAddress,"常総市") then
'	if inStr(strAddress,"菅生町") then
'	if inStr(strAddress,"下野原") then
'	if inStr(strAddress,"４７７") then
'	if inStr(strAddress,"１") then
'		GetHaitatsu = "茨城県 常総市菅生町下野原477-1 日本通運㈱ 守谷センター<br>支店止め又はチャーター"
'		exit function
'	end if
'	end if
'	end if
'	end if
'	end if
	if inStr(strAddress,"那珂郡") then
	if inStr(strAddress,"東海村") then
	if inStr(strAddress,"照沼") then
	if inStr(strAddress,"７６８") then
	if inStr(strAddress,"２３") then
		GetHaitatsu = "茨城県 那珂郡東海村照沼768-23 常陸那珂火力発電所<br>支店止め又はチャーター"
		exit function
	end if
	end if
	end if
	end if
	end if
	if inStr(strAddress,"結城郡") then
	if inStr(strAddress,"八千代町") then
	if inStr(strAddress,"平塚") then
	if inStr(strAddress,"菱毛道西") then
	if inStr(strAddress,"４４４８") then
		GetHaitatsu = "茨城県 結城郡八千代町平塚字菱毛道西4448 エフピコ物流関東ハブセンター<br>支店止め又はチャーター"
		exit function
	end if
	end if
	end if
	end if
	end if
	'-----------------------------------------------------------------------
	'栃木県
	'-----------------------------------------------------------------------
	if inStr(strAddress,"日光市") then
	if inStr(strAddress,"五十里") _
	or inStr(strAddress,"上栗山") _
	or inStr(strAddress,"上三依") _
	or inStr(strAddress,"川俣") _
	or inStr(strAddress,"川治温泉") _
	or inStr(strAddress,"黒部") _
	or inStr(strAddress,"鶏頂山") _
	or inStr(strAddress,"芹沢") _
	or inStr(strAddress,"高原") _
	or inStr(strAddress,"高原") _
	or inStr(strAddress,"独鈷沢") _
	or inStr(strAddress,"土呂部") _
	or inStr(strAddress,"中三依") _
	or inStr(strAddress,"西川") _
	or inStr(strAddress,"野門") _
	or inStr(strAddress,"日蔭") _
	or inStr(strAddress,"日向") _
	or inStr(strAddress,"湯西川") _
	or inStr(strAddress,"横川") _
	or inStr(strAddress,"若間") _
								  then
		GetHaitatsu = "栃木県 日光市（五十里、上栗山、上三依、川俣、川治温泉、黒部、鶏頂山、芹沢、高原、独鈷沢、土呂部、中三依、西川、野門、日蔭、日向、藤原1000番台、湯西川、横川、若間）<br>配達不能地区"
		exit function
	end if
	end if
	if inStr(strAddress,"日光市") then
	if inStr(strAddress,"藤原") then
	if inStr1000(strAddress) then
		GetHaitatsu = "栃木県 日光市（五十里、上栗山、上三依、川俣、川治温泉、黒部、鶏頂山、芹沢、高原、独鈷沢、土呂部、中三依、西川、野門、日蔭、日向、藤原1000番台、湯西川、横川、若間）<br>配達不能地区"
		exit function
	end if
	end if
	end if
	if inStr(strAddress,"日光市") then
	if inStr(strAddress,"中宮詞") _
	or inStr(strAddress,"湯元") then
		GetHaitatsu = "栃木県 日光市中宮詞、湯元<br>週2回（火・金）配達"
		exit function
	end if
	end if
	'-----------------------------------------------------------------------
	'群馬県
	'-----------------------------------------------------------------------
	if inStr(strAddress,"桐生市") then
	if inStr(strAddress,"黒保根町") then
	if inStr(strAddress,"みどり市東町") then
		GetHaitatsu = "群馬県 桐生市黒保根町 みどり市東町<br>週2回（火・金）配達"
		exit function
	end if
	end if
	end if
	if inStr(strAddress,"吾妻郡") then
	if inStr(strAddress,"草津町") then
		GetHaitatsu = "群馬県 吾妻郡草津町<br>週3回（月・水・金）配達"
		exit function
	end if
	end if
	if inStr(strAddress,"吾妻郡") then
	if inStr(strAddress,"嬬恋村") then
		GetHaitatsu = "群馬県 吾妻郡嬬恋村・吾妻郡長野原町北軽井沢<br>週3回（火・木・土）配達"
		exit function
	end if
	end if
	if inStr(strAddress,"吾妻郡") then
	if inStr(strAddress,"長野原町") then
	if inStr(strAddress,"北軽井沢") then
		GetHaitatsu = "群馬県 吾妻郡嬬恋村・吾妻郡長野原町北軽井沢<br>週3回（火・木・土）配達"
		exit function
	end if
	end if
	end if
	if inStr(strAddress,"吾妻郡") then
	if inStr(strAddress,"中之条町") then
	if inStr(strAddress,"赤岩") _
	or inStr(strAddress,"入山") _
	or inStr(strAddress,"太子") _
	or inStr(strAddress,"小雨") _
	or inStr(strAddress,"生須") _
	or inStr(strAddress,"日影") _
									then
		GetHaitatsu = "群馬県 吾妻郡中之条町（赤岩・入山・太子・小雨・生須・日影）・前橋市富士見町赤城山<br>週1回月曜日配達"
		exit function
	end if
	end if
	end if
	if inStr(strAddress,"前橋市") then
	if inStr(strAddress,"富士見町") then
	if inStr(strAddress,"赤城山") then
		GetHaitatsu = "群馬県 吾妻郡中之条町（赤岩・入山・太子・小雨・生須・日影）・前橋市富士見町赤城山<br>週1回月曜日配達"
		exit function
	end if
	end if
	end if
	'埼玉県
	GetHaitatsu = Saitama(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'千葉県
	'-----------------------------------------------------------------------
	GetHaitatsu = Chiba(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'神奈川
	'-----------------------------------------------------------------------
	GetHaitatsu = Kanagawa(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'東京都
	'-----------------------------------------------------------------------
	GetHaitatsu = Tokyo(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'山梨県
	'-----------------------------------------------------------------------
	GetHaitatsu = Yamanashi(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'新潟県
	'-----------------------------------------------------------------------
	GetHaitatsu = Nigata(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'長野県
	'-----------------------------------------------------------------------
	GetHaitatsu = Nagano(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'富山県
	'-----------------------------------------------------------------------
	GetHaitatsu = Toyama(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'福井県
	'-----------------------------------------------------------------------
	GetHaitatsu = Fukui(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'静岡県
	'-----------------------------------------------------------------------
	GetHaitatsu = Shizuoka(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'岐阜県
	'-----------------------------------------------------------------------
	GetHaitatsu = Gifu(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'愛知県
	'-----------------------------------------------------------------------
	GetHaitatsu = Aichi(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'三重県
	'-----------------------------------------------------------------------
	GetHaitatsu = Mie(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'滋賀県
	'-----------------------------------------------------------------------
	GetHaitatsu = Shiga(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'京都府
	'-----------------------------------------------------------------------
	GetHaitatsu = Kyoto(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'和歌山県
	'-----------------------------------------------------------------------
	GetHaitatsu = Wakayama(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'大阪府
	GetHaitatsu = Osaka(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'兵庫県
	'-----------------------------------------------------------------------
	GetHaitatsu = Hyogo(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'奈良県
	'-----------------------------------------------------------------------
	GetHaitatsu = Nara(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'岡山県
	'-----------------------------------------------------------------------
	GetHaitatsu = Okayama(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'広島県
	'-----------------------------------------------------------------------
	GetHaitatsu = Hiroshima(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'鳥取県
	'-----------------------------------------------------------------------
	GetHaitatsu = Tottori(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'島根県
	'-----------------------------------------------------------------------
	GetHaitatsu = Shimane(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'山口県
	'-----------------------------------------------------------------------
	GetHaitatsu = Yamaguchi(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'徳島県
	'-----------------------------------------------------------------------
	GetHaitatsu = Tokushima(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'香川県
	'-----------------------------------------------------------------------
	GetHaitatsu = Kagawa(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'愛媛県
	'-----------------------------------------------------------------------
	GetHaitatsu = Ehime(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'-----------------------------------------------------------------------
	'高知県
	'-----------------------------------------------------------------------
	GetHaitatsu = Kochi(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'熊本県
	GetHaitatsu = Kumamoto(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'大分県
	GetHaitatsu = Oita(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'宮崎
	GetHaitatsu = Miyazaki(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
	'沖縄県
	GetHaitatsu = Okinawa(strAddress)
	if GetHaitatsu <> "" then
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'山梨県
'2019.01.10
'-----------------------------------------------------------------------
Private Function Yamanashi(byVal strAddress)
	Yamanashi = ""
	if inAddress(strAddress,"山梨") = False then
		exit function
	end if
	if inAddressEx(strAddress,"南巨摩郡.*早川町") then
'	or inAddressEx(strAddress,"中巨摩郡.*昭和町") then
		Yamanashi = "山梨県 南巨摩郡早川町" _
				 & "<br>週２回配達"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'新潟県
'2019.01.10
'-----------------------------------------------------------------------
Private Function Nigata(byVal strAddress)
	Nigata = ""
	if inAddress(strAddress,"新潟") = False then
		exit function
	end if
	if inAddressEx(strAddress,"柏崎市.*青山町") then
'	or inAddressEx(strAddress,"柏崎市.*藤井") then
		Nigata = "新潟県 柏崎市青山町 柏崎原子力発電所 及びその関連企業" _
				 & "<br>配達不可（チャーターも不可）"
		exit function
	end if
	if inAddressEx(strAddress,"柏崎市.*高柳町") then
'	or inAddressEx(strAddress,"柏崎市.*藤井") then
		Nigata = "新潟県 柏崎市高柳町" _
				 & "<br>冬期配達不可（12月1日～3月31日）"
		exit function
	end if
	if inAddress(strAddress,"長岡市") then
		if inAddressEx(strAddress,"長岡市.*蓬平町") _
		or inAddressEx(strAddress,"長岡市.*乙吉町") _
		or inAddressEx(strAddress,"長岡市.*濁沢町") _
		or inAddressEx(strAddress,"長岡市.*水穴町") _
		or inAddressEx(strAddress,"長岡市.*竹之高地町") _
		or inAddressEx(strAddress,"長岡市.*軽井沢") _
		or inAddressEx(strAddress,"長岡市.*西中野俣") _
		or inAddressEx(strAddress,"長岡市.*東中野俣") _
		or inAddressEx(strAddress,"長岡市.*半蔵金") _
		or inAddressEx(strAddress,"長岡市.*田代") _
		or inAddressEx(strAddress,"長岡市.*森上") _
		or inAddressEx(strAddress,"長岡市.*山葵谷") _
		or inAddressEx(strAddress,"長岡市.*葎谷") then
	'	or inAddressEx(strAddress,"長岡市.*新産") then
			Nigata = "新潟県 長岡市（蓬平町、乙吉町、濁沢長、水穴町、竹之高地町、軽井沢、西中野俣、東中野俣、半蔵金、田代、森上、山葵谷、葎谷）" _
					 & "<br>冬期配達不可（12月1日～3月31日）"
		end if
		exit function
	end if
	if inAddress(strAddress,"三条市") then
		if inAddressEx(strAddress,"三条市.*名下") _
		or inAddressEx(strAddress,"三条市.*大谷地") _
		or inAddressEx(strAddress,"三条市.*笠堀") then
'		or inAddressEx(strAddress,"三条市.*南四日町") then
			Nigata = "新潟県 三条市（名下、大谷地、笠堀）" _
					 & "<br>冬期配達不可（12月1日～3月31日）"
		end if
		exit function
	end if
	if inAddress(strAddress,"小千谷市") then
		if inAddressEx(strAddress,"小千谷市.*小栗山") _
		or inAddressEx(strAddress,"小千谷市.*南荷頃") _
		or inAddressEx(strAddress,"小千谷市.*塩谷") _
		or inAddressEx(strAddress,"小千谷市.*真人町") _
		or inAddressEx(strAddress,"小千谷市.*池中新田") _
		or inAddressEx(strAddress,"小千谷市.*岩沢") _
		or inAddressEx(strAddress,"小千谷市.*川井") then
'		or inAddressEx(strAddress,"魚沼市.*山田") then
			Nigata = "新潟県 小千谷市（大字小栗山、大字南荷頃、大字塩谷、大字真人町、大字池中新田、大字岩沢、大字川井）" _
					 & "<br>冬期配達不可（12月1日～3月31日）"
		end if
		exit function
	end if
	if inAddress(strAddress,"加茂市") then
		if inAddressEx(strAddress,"加茂市.*宮寄上") _
		or inAddressEx(strAddress,"加茂市.*上高柳") _
		or inAddressEx(strAddress,"加茂市.*下高柳") _
		or inAddressEx(strAddress,"加茂市.*西山") _
		or inAddressEx(strAddress,"加茂市.*中大谷") _
		or inAddressEx(strAddress,"加茂市.*下大谷") _
		or inAddressEx(strAddress,"加茂市.*上大谷") _
		or inAddressEx(strAddress,"加茂市.*黒水") _
		or inAddressEx(strAddress,"加茂市.*上土倉") _
		or inAddressEx(strAddress,"加茂市.*下土倉") _
		or inAddressEx(strAddress,"加茂市.*長谷") _
		or inAddressEx(strAddress,"加茂市.*狭口") then
'		or inAddressEx(strAddress,"柏崎市.*松美") then
			Nigata = "新潟県 加茂市（大字宮寄上、大字上高柳、大字下高柳、大字西山、大字中大谷、" _
				   & "大字下大谷、大字上大谷、大字黒水、大字上土倉、大字下土倉、長谷、狭口）" _
				   & "<br>冬期配達不可（12月1日～3月31日）"
		end if
		exit function
	end if
	if inAddress(strAddress,"南魚沼") then
		if inAddressEx(strAddress,"南魚沼.*石打") then
'		or inAddressEx(strAddress,"南魚沼.*六日町") then
			Nigata = "新潟県 南魚沼市石打 石打丸山スキー場" _
				   & "<br>冬期配達不能（12月1日～4月30日）"
		elseif inAddressEx(strAddress,"南魚沼.*湯沢町.*土樽") then
'			or inAddressEx(strAddress,"南魚沼.*六日町") then
			Nigata = "新潟県 南魚沼郡湯沢町土樽731 岩原スキー場" _
				   & "<br>冬期配達不能（12月1日～4月30日）"
		end if
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'長野県
'2019.01.14
'-----------------------------------------------------------------------
Private Function Nagano(byVal strAddress)
	Nagano = ""
	if inAddress(strAddress,"長野") = False then
		exit function
	end if
	if inAddressEx(strAddress,"大町市.*平") then
'	or inAddressEx(strAddress,"大町市.*大町") then
		if inAddressEx(strAddress,"大町市.*平.*２") then
'		or inAddressEx(strAddress,"大町市.*大町.*１") then
			if inAddressEx(strAddress,"大町市.*平.*２１１７") then
'			or inAddressEx(strAddress,"大町市.*大町.*１５４７") then
				Nagano = "長野県 大町市平20000～（鹿島槍スキー場付近）" _
					 & "<br>冬期配達不可（11月～4月）"
			else
				Nagano = "長野県 大町市平2117" _
					 & "<br>週2回（月・木）のみ配達"
			end if
			exit function
		end if
	end if
	if inAddressEx(strAddress,"松本市.*安曇.*上高地") then
'	or inAddressEx(strAddress,"松本市.*新村.*２３７０") then
		Nagano = "長野県 松本市安曇上高地" _
			   & "<br>冬期配達不可（11月～4月）"
	end if
	if inAddressEx(strAddress,"下高井郡") then
'	or inAddressEx(strAddress,"佐久市") then
		if inAddressEx(strAddress,"下高井郡.*山.*内町.*夜間瀬") then
'		or inAddressEx(strAddress,"佐久市.*岩村田") then
			Nagano = "長野県 下高井郡山ノ内町夜間瀬11700-○○、12347-○○ 竜王スキー場内" _
				   & "<br>冬期配達不可（12月1日～4月30日）"
			exit function
		elseif inAddressEx(strAddress,"下高井郡.*山.*内町.*平穏") then
'		or inAddressEx(strAddress,"佐久市.*猿久保") then
			Nagano = "長野県 下高井郡山ノ内町大字平穏7148-○○、7149-○○（志賀高原）" _
				   & "<br>冬期配達不可（12月1日～4月30日）"
			exit function
		elseif inAddressEx(strAddress,"下高井郡.*野沢温泉") then
'		or inAddressEx(strAddress,"佐久市.*長土呂.*下北原") then
			Nagano = "長野県 下高井郡野沢温泉村大字豊郷8003～8455番地 野沢温泉スキー場内" _
				   & "<br>冬期配達不可（12月1日～4月30日）"
			exit function
		end if
	end if
	if inAddressEx(strAddress,"北安曇郡") then
'	or inAddressEx(strAddress,"安曇") then
		Nagano = "長野県 北安曇郡" _
			   & "<br>冬期配達不可（12月1日～4月30日）"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'富山県
'2019.01.14
'-----------------------------------------------------------------------
Private Function Toyama(byVal strAddress)
	Toyama = ""
	if inAddress(strAddress,"富山") = False then
		exit function
	end if
	if inAddress(strAddress,"富山市") then
		if inAddress(strAddress,"富山市有峰") _
		or inAddress(strAddress,"富山市粟巣野") _
		or inAddress(strAddress,"富山市大山松木") _
		or inAddress(strAddress,"富山市岡田") _
		or inAddress(strAddress,"富山市小見") _
		or inAddress(strAddress,"富山市亀谷") _
		or inAddress(strAddress,"富山市才覚地") _
		or inAddress(strAddress,"富山市中地山") _
		or inAddress(strAddress,"富山市原") _
		or inAddress(strAddress,"富山市本宮") _
		or inAddress(strAddress,"富山市牧") _
		or inAddress(strAddress,"富山市水須") _
		or inAddress(strAddress,"富山市和田") _
		or inAddress(strAddress,"富山市山田") _
		or inAddress(strAddress,"富山市八尾町赤石") _
		or inAddress(strAddress,"富山市八尾町新谷") _
		or inAddress(strAddress,"富山市八尾町庵谷") _
		or inAddress(strAddress,"富山市八尾町内名") _
		or inAddress(strAddress,"富山市八尾町大玉生") _
		or inAddress(strAddress,"富山市八尾町尾畑") _
		or inAddress(strAddress,"富山市八尾町桂原") _
		or inAddress(strAddress,"富山市八尾町上牧") _
		or inAddress(strAddress,"富山市八尾町桐谷") _
		or inAddress(strAddress,"富山市八尾町栗須") _
		or inAddress(strAddress,"富山市八尾町小井波") _
		or inAddress(strAddress,"富山市八尾町島地") _
		or inAddress(strAddress,"富山市八尾町清水") _
		or inAddress(strAddress,"富山市八尾町下島") _
		or inAddress(strAddress,"富山市八尾町杉平") _
		or inAddress(strAddress,"富山市八尾町薄尾") _
		or inAddress(strAddress,"富山市八尾町外堀") _
		or inAddress(strAddress,"富山市八尾町高野") _
		or inAddress(strAddress,"富山市八尾町谷折") _
		or inAddress(strAddress,"富山市八尾町田頭") _
		or inAddress(strAddress,"富山市八尾町栃折") _
		or inAddress(strAddress,"富山市八尾町中島") _
		or inAddress(strAddress,"富山市八尾町中山") _
		or inAddress(strAddress,"富山市八尾町西原") _
		or inAddress(strAddress,"富山市八尾町花房") _
		or inAddress(strAddress,"富山市八尾町東原") _
		or inAddress(strAddress,"富山市八尾町東松瀬") _
		or inAddress(strAddress,"富山市八尾町水無") _
		then
'		or inAddress(strAddress,"富山市才覚寺") then
			Toyama = "富山県 富山市（有峰、粟巣野、大山松木、岡田、小見、亀谷、才覚地、中地山、原、本宮、牧、水須、和田、山田○○、" _
				   & "八尾町（赤石、新谷、庵谷、内名、大玉生、尾畑、桂原、上牧、桐谷、栗須、小井波、島地、清水、下島、杉平、" _
				   & "薄尾、外堀、高野、谷折、田頭、栃折、中島、中山、西原、花房、東原、東松瀬、水無）" _
				   & "<br>冬期配達不可（12月～3月）"
		end if
		exit function
	end if
	if inAddress(strAddress,"魚津市虎谷") then
'	or inAddress(strAddress,"魚津市大光寺") then
		Toyama = "富山県 魚津市虎谷<br>冬期配達不可（12月～3月）"
	end if
	if inAddress(strAddress,"中新川郡上市町伊折") _
	or inAddress(strAddress,"中新川郡上市町稲村") _
	or inAddress(strAddress,"中新川郡上市町折戸") _
	or inAddress(strAddress,"中新川郡上市町千石") _
	or inAddress(strAddress,"中新川郡上市町下田") _
	or inAddress(strAddress,"中新川郡上市町中村") _
	or inAddress(strAddress,"中新川郡上市町西種") _
	or inAddress(strAddress,"中新川郡上市町東種") _
	or inAddress(strAddress,"中新川郡上市町蓬沢") _
	or inAddress(strAddress,"中新川郡立山町芦峅寺") _
	or inAddress(strAddress,"中新川郡立山町伊勢屋") _
	or inAddress(strAddress,"中新川郡立山町小又") _
	or inAddress(strAddress,"中新川郡立山町谷") _
	or inAddress(strAddress,"中新川郡立山町千垣") _
	or inAddress(strAddress,"中新川郡立山町松倉") _
	or inAddress(strAddress,"中新川郡立山町目桑") then
'	or inAddress(strAddress,"中新川郡上市町上経田") then
		Toyama = "富山県 中新川郡上市町（伊折、稲村、折戸、千石、下田、中村、西種、東種、蓬沢）、" _
			   & "立山町（芦峅寺（千寿ヶ原、ブナ坂）、伊勢屋、小又、谷、千垣、松倉、目桑）" _
			   & "<br>冬期配達不可（12月～3月）"
		exit function
	end if
	if inAddress(strAddress,"南砺市") then
		if inAddress(strAddress,"南砺市利賀村") _
		or inAddress(strAddress,"南砺市平村") _
		or inAddress(strAddress,"南砺市相倉") _
		or inAddress(strAddress,"南砺市入谷") _
		or inAddress(strAddress,"南砺市大崩島") _
		or inAddress(strAddress,"南砺市大島") _
		or inAddress(strAddress,"南砺市篭渡") _
		or inAddress(strAddress,"南砺市上梨") _
		or inAddress(strAddress,"南砺市上松尾") _
		or inAddress(strAddress,"南砺市来栖") _
		or inAddress(strAddress,"南砺市小来栖") _
		or inAddress(strAddress,"南砺市下出") _
		or inAddress(strAddress,"南砺市下梨") _
		or inAddress(strAddress,"南砺市寿川") _
		or inAddress(strAddress,"南砺市杉尾") _
		or inAddress(strAddress,"南砺市祖山") _
		or inAddress(strAddress,"南砺市高草嶺") _
		or inAddress(strAddress,"南砺市田向") _
		or inAddress(strAddress,"南砺市渡原") _
		or inAddress(strAddress,"南砺市中畑") _
		or inAddress(strAddress,"南砺市梨谷") _
		or inAddress(strAddress,"南砺市夏焼") _
		or inAddress(strAddress,"南砺市東中江") _
		or inAddress(strAddress,"南砺市見座") _
		or inAddress(strAddress,"南砺市上平村") _
		or inAddress(strAddress,"南砺市新屋") _
		or inAddress(strAddress,"南砺市猪谷") _
		or inAddress(strAddress,"南砺市打越") _
		or inAddress(strAddress,"南砺市漆谷") _
		or inAddress(strAddress,"南砺市小瀬") _
		or inAddress(strAddress,"南砺市小原") _
		or inAddress(strAddress,"南砺市皆葎") _
		or inAddress(strAddress,"南砺市桂") _
		or inAddress(strAddress,"南砺市上平細島") _
		or inAddress(strAddress,"南砺市上中田") _
		or inAddress(strAddress,"南砺市楮") _
		or inAddress(strAddress,"南砺市下島") _
		or inAddress(strAddress,"南砺市菅沼") _
		or inAddress(strAddress,"南砺市田下") _
		or inAddress(strAddress,"南砺市成出") _
		or inAddress(strAddress,"南砺市西赤尾町") _
		or inAddress(strAddress,"南砺市東赤尾") _
		or inAddress(strAddress,"南砺市真木") _
		or inAddress(strAddress,"南砺市葎島") _
		or inAddress(strAddress,"南砺市井波") then
			Toyama = "富山県 南砺市（旧利賀村）利賀村○○・（旧平村）相倉、入谷、大崩島、大島、篭渡、" _
				   & "上梨、上松尾、来栖、小来栖、下出、下梨、寿川、杉尾、祖山、高草嶺、田向、渡原、" _
				   & "中畑、梨谷、夏焼、東中江、見座・（旧上平村）新屋、猪谷、打越、漆谷、小瀬、小原、" _
				   & "皆葎、桂、上平細島、上中田、楮、下島、菅沼、田下、成出、西赤尾町、東赤尾、真木、葎島" _
				   & "<br>【週１～２回配達（4月～11月）※軽四以上はチャーター】" _
				   & "<br>冬季配達不能（12月～3月）"
		end if
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'福井県
'2019.01.15
'-----------------------------------------------------------------------
Private Function Fukui(byVal strAddress)
	Fukui = ""
	if inAddress(strAddress,"福井") = False then
		exit function
	end if
	if inAddress(strAddress,"発電所") then		' 、岩神
		if inAddress(strAddress,"三方郡美浜町") _
		or inAddress(strAddress,"大飯郡大飯町") _
		or inAddress(strAddress,"大飯郡高浜町") then
			Fukui = "福井県 三方郡美浜町美浜原子力発電所、大飯郡大飯町大島大飯発電所、高浜町田ノ浦高浜発電所" _
				  & "<br>チャータ扱い"
			exit function
		end if
	end if
	if inAddress(strAddress,"敦賀市杉津") then
		if inAddress(strAddress,"パーキング") then
			Fukui = "福井県 敦賀市杉津　杉津パーキングエリア宛" _
				  & "<br>冬季配達不可（12月～2月）"
			exit function
		end if
	end if
	if inAddress(strAddress,"敦賀市白木１") _
	or inAddress(strAddress,"敦賀市白木２") _
	or inAddress(strAddress,"敦賀市明神町") then
		Fukui = "福井県 敦賀市(白木1丁目、2丁目・明神町)" _
			  & "<br>チャータ扱い"
		exit function
	end if
	if inAddress(strAddress,"敦賀市浦底") _
	or inAddress(strAddress,"敦賀市立石") then
		Fukui = "福井県 敦賀市（浦底、立石)" _
			  & "<br>チャータ扱い"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'静岡県
'2019.01.20
'-----------------------------------------------------------------------
Private Function Shizuoka(byVal strAddress)
	Shizuoka = ""
	if inAddress(strAddress,"静岡") = False then
		exit function
	end if
	if inAddress(strAddress,"三島市山田新田４９") then
'	or inAddress(strAddress,"御前崎市池新田３７")	then
		Shizuoka = "静岡県 三島市山田新田４９０８芦の湖高原" _
				 & "<br>配達不能地区"
	end if
	if inAddress(strAddress,"三島市山田新田４７") then
'	or inAddress(strAddress,"御前崎市池新田３７")	then
		Shizuoka = "静岡県 三島市山田新田４７００番台 芦ノ湖高原別荘地" _
				 & "<br>配達不能地区"
	end if
	if inAddress(strAddress,"掛川市光陽２０８") then
'	or inAddress(strAddress,"掛川市天王町７")	then
		Shizuoka = "静岡県 掛川市光陽208 加藤産業株式会社" _
				 & "<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"裾野市茶畑") then
'	or inAddress(strAddress,"御殿場市新橋字堀向")	then
		Shizuoka = "静岡県 裾野市茶畑2250-1 芦ノ湖スカイラインレストハウスフジビュー" _
				 & "<br>配達不能地区"
	end if

	if inAddress(strAddress,"静岡市葵区") then
		if inAddress(strAddress,"静岡市葵区井川") _
		or inAddress(strAddress,"静岡市葵区岩崎") _
		or inAddress(strAddress,"静岡市葵区大沢") _
		or inAddressEx(strAddress,"静岡市葵区奥池.*谷") _
		or inAddress(strAddress,"静岡市葵区奥仙俣") _
		or inAddress(strAddress,"静岡市葵区落合柿島") _
		or inAddress(strAddress,"静岡市葵区上落合") _
		or inAddress(strAddress,"静岡市葵区上坂本") _
		or inAddress(strAddress,"静岡市葵区口坂本") _
		or inAddress(strAddress,"静岡市葵区口仙俣") _
		or inAddress(strAddress,"静岡市葵区小河内") _
		or inAddress(strAddress,"静岡市葵区腰越") _
		or inAddress(strAddress,"静岡市葵区内匠") _
		or inAddress(strAddress,"静岡市葵区長熊") _
		or inAddress(strAddress,"静岡市葵区長妻田森腰") _
		or inAddress(strAddress,"静岡市葵区油野") _
		or inAddress(strAddress,"静岡市葵区横沢") then
'		or inAddress(strAddress,"静岡市葵区上土") then
			Shizuoka = "静岡県 静岡市葵区（井川、岩崎、大沢、奥池ヶ谷、奥仙俣 落合柿島、上落合、上坂本、口坂本、口仙俣、" _
	                 & "小河内、腰越 内匠、長熊、長妻田森腰 油野、横沢）" _
					 & "<br>週３回（月・水・金）配達"
		elseif inAddress(strAddress,"静岡市葵区田代") _
			or inAddress(strAddress,"静岡市葵区井川") then
			Shizuoka = "静岡県 静岡市葵区（田代○○　及び　井川２６２９－１９０　リバウェル井川スキー場）" _
					 & "<br>１ｔ以上はチャーター（但し週3回配達（月・水・金）"
		elseif inAddress(strAddress,"静岡市葵区有東木") _
			or inAddress(strAddress,"静岡市葵区梅ヶ島") _
			or inAddress(strAddress,"静岡市葵区渡中平") _
			or inAddress(strAddress,"静岡市葵区入島") _
			or inAddress(strAddress,"静岡市葵区平野横山") _
			or inAddress(strAddress,"静岡市葵区蕨野") then
'			or inAddress(strAddress,"静岡市葵区上土") then
			Shizuoka = "静岡県 静岡市葵区（有東木、梅ヶ島 、渡中平、入島  平野横山 蕨野）" _
					 & "<br>週３回（火・木・土）配達"
		end if
		exit function
	end if
end function
'-----------------------------------------------------------------------
'岐阜県
'2019.01.21
'-----------------------------------------------------------------------
Private Function Gifu(byVal strAddress)
	Gifu = ""
	if inAddress(strAddress,"岐阜") = False then
		exit function
	end if
	if inAddress(strAddress,"揖斐郡揖斐川町") then	' 大垣市外渕
		if inAddress(strAddress,"開田、門入、塚、徳山、戸入、櫨原、東杉原、鶴見")	then
			Gifu = "岐阜県 揖斐郡揖斐川町（開田、門入、塚、徳山、戸入、櫨原、東杉原、鶴見）" _
				 & "<br>配達不能地区"
		elseif inAddress(strAddress,"春日、小津、乙原、樫原、坂内、谷汲、外津汲、西津汲、西横山、東津汲、東横山、日坂、三倉、樒平、下平、椿井野")	then
			Gifu = "岐阜県 揖斐郡揖斐川町（春日○○、小津、乙原、樫原、坂内○○、谷汲○○、外津汲、西津汲、" _
				 & "西横山、東津汲、東横山、日坂、三倉、樒平、下平、椿井野）" _
				 & "<br>週1回木曜日配達"
		end if
		exit function
	end if
	if inAddress(strAddress,"下呂市")	then
	if inAddress(strAddress,"小坂町")	then
	if inAddress(strAddress,"落合濁河温泉")	then
		Gifu = "岐阜県 下呂市小坂町落合濁河温泉<br>配達不能地区"
		exit function
	end if
	end if
	end if

	if inAddress(strAddress,"本巣市")	then
	if inAddress(strAddress,"金原、神海、木知原、佐原、曽井中島、外山、日当、法林寺、文殊、山口、根尾")	then
		Gifu = "岐阜県 本巣市(金原､神海､木知原､佐原､曽井中島､外山､日当､法林寺､文殊､山口､根尾○○)<br>週２回配達（火・木のみ）"
		exit function
	end if
	end if

	if inAddress(strAddress,"高山市")	then
	if inAddress(strAddress,"朝日町、荘川町、高根町、上宝町、奥飛騨温泉郷")	then
		Gifu = "岐阜県 高山市朝日町○○、荘川町○○、高根町○○、上宝町○○、奥飛騨温泉郷○○<br>週２回配達"
		exit function
	end if
	end if
	if inAddress(strAddress,"高山市")	then
	if inAddress(strAddress,"岩井町")	_
	or (inAddress(strAddress,"清見町") and inAddress(strAddress,"大原、楢谷")) _
	or (inAddress(strAddress,"高根町") and inAddress(strAddress,"日和田、小日和田、留之原、野麦、黍生、阿多野郷")) _
										then
		Gifu = "岐阜県 高山市岩井町、清見町(大原・楢谷）、高根町（日和田、小日和田、留之原、野麦、黍生、阿多野郷）<br>冬季配達不能（12～3月）"
		exit function
	end if
	end if
	if inAddress(strAddress,"高山市")	then
	if inAddress(strAddress,"上宝町")	then
	if inAddress(strAddress,"蔵柱")		then
	if inAddress(strAddress,"天文台")	then
		Gifu = "岐阜県 高山市上宝町蔵柱 京都大学飛騨天文台<br>配達不能地区"
		exit function
	end if
	end if
	end if
	end if

	if inAddress(strAddress,"土岐市")	then
	if inAddress(strAddress,"土岐ヶ丘１")	then
		Gifu = "岐阜県 土岐ヶ丘1-2 プレミアムアウトレット各店舗<br>支店止め又はチャーター"
		exit function
	end if
	end if

	if inAddress(strAddress,"飛騨市")	then
	if inAddress(strAddress,"神岡町")	then
	if inAddress(strAddress,"森茂、和佐府、伊西、瀬戸、佐古、岩井谷、和佐保、打保、下之本")		then
		Gifu = "岐阜県 飛騨市神岡町（森茂、和佐府、伊西、瀬戸、佐古、岩井谷、和佐保、打保、下之本）<br>冬季配達不能（12～3月）"
		exit function
	end if
	end if
	end if
	if inAddress(strAddress,"飛騨市")	then
	if inAddress(strAddress,"河合町、宮川町")	then
		Gifu = "岐阜県 飛騨市（河合町、宮川町）<br>冬季週１回（金のみ）配達(12～3月)"
		exit function
	end if
	end if
	if inAddress(strAddress,"恵那市")	then
	if inAddress(strAddress,"武並町、中野方町、東野、三郷町")	then
		Gifu = "岐阜県 恵那市（武並町○○、中野方町、東野、三郷町○○）<br>週２回（月・木）配達"
		exit function
	end if
	end if
	if inAddress(strAddress,"恵那市")	then
	if inAddress(strAddress,"飯地町")	then
		Gifu = "岐阜県 恵那市飯地町<br>週２回（火・金）配達"
		exit function
	end if
	end if
	if inAddress(strAddress,"大野郡")	then
	if inAddress(strAddress,"白川村")	then
		Gifu = "岐阜県 大野郡白川村<br>週２回（月・木）配達"
		exit function
	end if
	end if
	if inAddress(strAddress,"瑞浪市、土岐市、多治見市")	then
		Gifu = "岐阜県 瑞浪市・土岐市・多治見市" _
			 & "<br>現場宛ての1件1個の実重量が45kg以上、" _
			 & "または3辺合計280cm以上の商品及び1件100kg以上の商品は、支店止め又はチャーター"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'北海道
'2019.06.27
'-----------------------------------------------------------------------
Private Function Hokaido(byVal strAddress)
	Hokaido = ""
	'旭川市　江丹別町	週２回（火・金）のみ配達、時間指定不可
	if inAddress(strAddress,"旭川市")	then
	if inAddress(strAddress,"江丹別町")	then
		Hokaido = "北海道 旭川市　江丹別町	週２回（火・金）のみ配達、時間指定不可"
		exit function
	end if
	end if
	'三笠市ほん別鳥井沢町　桂沢ダム	冬期配達不可（12月1日～3月31日）
	if inAddress(strAddress,"三笠市")	then
	if inAddress(strAddress,"ほん別鳥井沢町")	then
	if inAddress(strAddress,"桂沢ダム")	then
		Hokaido = "北海道 三笠市ほん別鳥井沢町　桂沢ダム	冬期配達不可（12月1日～3月31日）"
		exit function
	end if
	end if
	end if
	'奥尻郡奥尻町	週３回配達（水・金・土）
	if inAddress(strAddress,"奥尻郡")	then
	if inAddress(strAddress,"奥尻町")	then
		Hokaido = "北海道 奥尻郡奥尻町	週３回配達（水・金・土）"
		exit function
	end if
	end if
	'上川郡上川町　層雲峡温泉	週２回（火・金）のみ配達、時間指定不可
	if inAddress(strAddress,"上川郡")	then
	if inAddress(strAddress,"上川町")	then
	if inAddress(strAddress,"層雲峡温泉")	then
		Hokaido = "北海道 上川郡上川町　層雲峡温泉	週２回（火・金）のみ配達、時間指定不可"
		exit function
	end if
	end if
	end if
	'上川郡東川町　旭岳温泉、天人峡温泉	週２回（火・金）のみ配達、時間指定不可
	if inAddress(strAddress,"上川郡")	then
	if inAddress(strAddress,"東川町")	then
	if inAddress(strAddress,"旭岳温泉、天人峡温")	then
		Hokaido = "北海道 上川郡東川町　旭岳温泉、天人峡温泉	週２回（火・金）のみ配達、時間指定不可"
		exit function
	end if
	end if
	end if
	'上川郡美瑛町　白金温泉	週２回（火・金）のみ配達、時間指定不可
	if inAddress(strAddress,"上川郡")	then
	if inAddress(strAddress,"美瑛町")	then
	if inAddress(strAddress,"白金温泉")	then
		Hokaido = "北海道 上川郡美瑛町　白金温泉	週２回（火・金）のみ配達、時間指定不可"
		exit function
	end if
	end if
	end if
	'空知郡上富良野町 十勝岳温泉	週２回（火・金）のみ配達、時間指定不可
	if inAddress(strAddress,"空知郡")	then
	if inAddress(strAddress,"上富良野町")	then
	if inAddress(strAddress,"十勝岳温泉")	then
		Hokaido = "北海道 空知郡上富良野町 十勝岳温泉	週２回（火・金）のみ配達、時間指定不可可"
		exit function
	end if
	end if
	end if
	'新冠郡新冠町 新冠ダム発電所	要事前相談
	if inAddress(strAddress,"新冠郡")	then
	if inAddress(strAddress,"新冠町")	then
	if inAddress(strAddress,"新冠ダム発電所")	then
		Hokaido = "北海道 新冠郡新冠町 新冠ダム発電所	要事前相談"
		exit function
	end if
	end if
	end if
	'日高郡新ひだか町の各ダム発電所	要事前相談
	if inAddress(strAddress,"日高郡")	then
	if inAddress(strAddress,"新ひだか町")	then
	if inAddress(strAddress,"ダム、発電所")	then
		Hokaido = "北海道 日高郡新ひだか町の各ダム発電所	要事前相談"
		exit function
	end if
	end if
	end if
	'河東郡上士幌 ぬかびら温泉郷	週１回土曜午後のみ配達
	if inAddress(strAddress,"河東郡")	then
	if inAddress(strAddress,"上士幌")	then
	if inAddress(strAddress,"ぬかびら温泉郷")	then
		Hokaido = "北海道 河東郡上士幌 ぬかびら温泉郷	週１回土曜午後のみ配達"
		exit function
	end if
	end if
	end if
	'河東郡鹿追町北瓜幕、河東郡鹿追町燃別湖畔	週１回土曜午後のみ配達
	if inAddress(strAddress,"河東郡")	then
	if inAddress(strAddress,"鹿追町")	then
	if inAddress(strAddress,"北瓜幕、燃別湖畔")	then
		Hokaido = "北海道 河東郡鹿追町北瓜幕、河東郡鹿追町燃別湖畔	週１回土曜午後のみ配達"
		exit function
	end if
	end if
	end if
	'中川郡豊頃町大津○○	週１回土曜午後のみ配達
	if inAddress(strAddress,"中川郡")	then
	if inAddress(strAddress,"豊頃町")	then
	if inAddress(strAddress,"大津")	then
		Hokaido = "北海道 中川郡豊頃町大津○○	週１回土曜午後のみ配達"
		exit function
	end if
	end if
	end if
	'TEST
'	if inAddress(strAddress,"苫小牧市")	then
'	if inAddress(strAddress,"三光町")	then
'		Hokaido = "北海道 TEST"
'		exit function
'	end if
'	end if
End Function
'-----------------------------------------------------------------------
'岩手県
'2019.06.28
'-----------------------------------------------------------------------
Private Function Iwate(byVal strAddress)
	Iwate = ""
	'岩手郡葛巻町葛巻36・43・44・45・51地割	冬期配達不可（12月1日～4月30日）
	if inAddress(strAddress,"葛巻町")	then
	if inAddress(strAddress,"葛巻")	then
	if inAddress(strAddress,"３６、４３、４４、４５、５１")	then
		Iwate = "岩手県 岩手郡葛巻町葛巻36・43・44・45・51地割	冬期配達不可（12月1日～4月30日）"
		exit function
	end if
	end if
	end if
End Function
'-----------------------------------------------------------------------
'秋田県
'2020.01.09
'-----------------------------------------------------------------------
Private Function Akita(byVal strAddress)
	Akita = ""
	if inAddress(strAddress,"秋田") = False then
		exit function
	end if
	if inAddressEx(strAddress,"大館市.*澄川.*１") then
'	or inAddressEx(strAddress,"大館市.*川口.*１") then
		Akita = "秋田県 大館市澄川1番地 三菱重工㈱田代試験場<br>冬期配達不可（12月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"大館市.*早口管谷地") then
'	or inAddressEx(strAddress,"大館市") then
		Akita = "秋田県 大館市早口管谷地34-2、㈱シムコ<br>冬期配達不可（12月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"鹿角市.*八幡平熊沢") _
	or inAddressEx(strAddress,"鹿角市.*八幡平切留平") _
	or inAddressEx(strAddress,"鹿角市.*八幡平湯瀬") then
'	or inAddressEx(strAddress,"大館市.*川口.*１") then
		Akita = "秋田県 鹿角市（八幡平熊沢○○、八幡平切留平・八幡平湯瀬十和田大湯字（熊取平、田代平、戸倉、西ノ森）" _
			  & "<br>冬期配達不可（12月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"仙北市.*田沢湖卒田夏瀬") then
'	or inAddressEx(strAddress,"横手市.*杉沢.*中杉沢") then
		Akita = "秋田県 仙北市田沢湖卒田夏瀬（夏瀬温泉都わすれホテル）" _
			  & "<br>冬期配達不可（12月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"仙北市.*田沢湖町.*玉川") then
'	or inAddressEx(strAddress,"横手市.*杉沢.*中杉沢") then
		Akita = "秋田県 仙北市田沢湖町玉川（玉川温泉・新玉川温泉・㈱ぶなの森)" _
			  & "<br>冬期配達不可（12月1日～4月15日）"
		exit function
	end if
	if inAddressEx(strAddress,"山本郡.*八峰町") _
	or inAddressEx(strAddress,"山本郡.*藤里町") then
'	or inAddressEx(strAddress,"山本郡.*三種町") then
		Akita = "秋田県 山本郡（八峰町峰浜水沢ダム・藤里町粕毛素波里ダム）" _
			  & "<br>冬期配達不可（12月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"鹿角郡.*小坂町.*十和田湖") then
'	or inAddressEx(strAddress,"山本郡.*三種町") then
		Akita = "秋田県 鹿角郡小坂町十和田湖" _
			  & "<br>冬期配達不可（12月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"雄勝郡.*羽後町.*軽井沢") _
	or inAddressEx(strAddress,"仙北郡.*美郷町.*飯詰") then
		Akita = "秋田県 雄勝郡羽後町軽井沢" _
			  & "<br>冬期配達不可（12月1日～3月31日）"
		exit function
	end if


End Function
'-----------------------------------------------------------------------
'山形県
'2018.03.14
'2019.01.09
'-----------------------------------------------------------------------
Private Function Yamagata(byVal strAddress)
	Yamagata = ""
	if inAddress(strAddress,"山形") = False then
		exit function
	end if
	if inAddressEx(strAddress,"酒田市.*飛島") _
	or inAddressEx(strAddress,"山形市.*面白山") _
	or inAddressEx(strAddress,"米沢市.*関.*白布") then
'	or inAddressEx(strAddress,"酒田市.*卸町") then
		Yamagata = "山形県 酒田市飛島（離島）山形市面白山 米沢市関字白布天元台" _
				 & "<br>配達不能地区"
		exit function
	end if
	if inAddressEx(strAddress,"西村山郡.*朝日町.*大暮山") _
	or inAddressEx(strAddress,"西村山郡.*朝日町.*大沼") _
	or inAddressEx(strAddress,"西村山郡.*朝日町.*水木") _
	or inAddressEx(strAddress,"西村山郡.*朝日町.*石須部") _
	or inAddressEx(strAddress,"西村山郡.*朝日町.*今平") _
	or inAddressEx(strAddress,"西村山郡.*朝日町.*立木") _
	or inAddressEx(strAddress,"西村山郡.*朝日町.*白倉") _
	or inAddressEx(strAddress,"西村山郡.*朝日町.*太郎") _
	or inAddressEx(strAddress,"西村山郡.*朝日町.*大船木") _
	or inAddressEx(strAddress,"西村山郡.*朝日町.*杉山") _
	or inAddressEx(strAddress,"西村山郡.*朝日町.*和合平") _
	or inAddressEx(strAddress,"西村山郡.*朝日町.*古槙") _
	or inAddressEx(strAddress,"西村山郡.*西川町.*大井沢")  _
	or inAddressEx(strAddress,"西村山郡.*西川町.*志津")  _
	or inAddressEx(strAddress,"西村山郡.*西川町.*岩根沢") then
'	or inAddressEx(strAddress,"西村山郡.*朝日町.*宮宿") then
		Yamagata = "山形県 西村山郡朝日町（大暮山・大沼・水木、石須部・今平・立木・白倉・太郎・大船木・杉山・和合平・古槙）、西川町（大井沢・志津・岩根沢）" _
				 & "<br>チャータ扱い"
		exit function
	end if
	if inAddressEx(strAddress,"上山市.*狸森") _
	or inAddressEx(strAddress,"上山市.*永野.*蔵王山") _
	or inAddressEx(strAddress,"上山市.*小倉") _
	or inAddressEx(strAddress,"上山市.*小白府") _
	or inAddressEx(strAddress,"上山市.*松山") _
	or inAddressEx(strAddress,"上山市.*鶴脛町") then
'	or inAddressEx(strAddress,"上山市") then
		Yamagata = "山形県 上山市（狸森、永野字蔵王山、小倉、小白府、松山、鶴脛町）" _
				 & "<br>冬期配達不可（11月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"寒河江市.*幸生") _
	or inAddressEx(strAddress,"寒河江市.*田代") then
'	or inAddressEx(strAddress,"寒河江市.*日田") then
		Yamagata = "山形県 寒河江市（幸生、大字田代）" _
				 & "<br>冬期配達不可（11月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"天童市.*田麦野") then
'	or inAddressEx(strAddress,"天童市.*乱川") then
		Yamagata = "山形県 天童市大字田麦野" _
				 & "<br>冬期配達不可（11月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"長井市.*館町北") then
'	or inAddressEx(strAddress,"長井市.*四.*谷") then
		Yamagata = "山形県 長井市館町北6-6（長井ダム）" _
				 & "<br>冬期配達不可（11月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"南陽市.*萩") _
	or inAddressEx(strAddress,"南陽市.*小滝") _
	or inAddressEx(strAddress,"南陽市.*釜渡戸") then
'	or inAddressEx(strAddress,"南陽市.*若狭郷屋.*沢見") then
		Yamagata = "山形県 南陽市（萩、小滝、釜渡戸）" _
				 & "<br>冬期配達不可（11月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"村山市.*五十沢") _
	or inAddressEx(strAddress,"村山市.*岩野") _
	or inAddressEx(strAddress,"村山市.*山ノ内") _
	or inAddressEx(strAddress,"村山市.*富並") _
	or inAddressEx(strAddress,"村山市.*樽石") _
	or inAddressEx(strAddress,"村山市.*土生田") _
	or inAddressEx(strAddress,"村山市.*湯野沢") _
	or inAddressEx(strAddress,"村山市.*白鳥") then
'	or inAddressEx(strAddress,"村山.*朝日") then
		Yamagata = "山形県 村山市（五十沢、岩野、山ノ内、富並、大字樽石、大字土生田、湯野沢、白鳥）" _
				 & "<br>冬期配達不可（11月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"山形市.*蔵王上野") _
	or inAddressEx(strAddress,"山形市.*蔵王堀田") _
	or inAddressEx(strAddress,"山形市.*蔵王温泉") _
	or inAddressEx(strAddress,"山形市.*山寺") _
	or inAddressEx(strAddress,"山形市.*関沢") _
	or inAddressEx(strAddress,"山形市.*八森") _
	or inAddressEx(strAddress,"山形市.*神尾") _
	or inAddressEx(strAddress,"山形市.*土坂") _
	or inAddressEx(strAddress,"山形市.*門伝") _
	or inAddressEx(strAddress,"山形市.*村木沢") then
		Yamagata = "山形県 山形市（蔵王上野、蔵王堀田、蔵王温泉、山寺、関沢、高沢、八森、神尾、土坂、門伝、村木沢）" _
				 & "<br>冬期配達不可（11月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"米沢市.*関") _
	or inAddressEx(strAddress,"米沢市.*大平") _
	or inAddressEx(strAddress,"米沢市.*大沢") _
	or inAddressEx(strAddress,"米沢市.*大小屋") _
	or inAddressEx(strAddress,"米沢市.*入田沢") _
	or inAddressEx(strAddress,"米沢市.*板谷") _
	or inAddressEx(strAddress,"米沢市.*栗谷") _
	or inAddressEx(strAddress,"米沢市.*網木") _
	or inAddressEx(strAddress,"米沢市.*小野川町") then
'	or inAddressEx(strAddress,"米沢市.*花沢") then
		Yamagata = "山形県 米沢市（関、大平、大沢、大小屋、入田沢、板谷、栗谷、網木、小野川町）" _
				 & "<br>冬期配達不可（11月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"西村山郡.*大江町.*柳川") _
	or inAddressEx(strAddress,"西村山郡.*大江町.*貫見") _
	or inAddressEx(strAddress,"西村山郡.*大江町.*月布") _
	or inAddressEx(strAddress,"西村山郡.*大江町.*十八才") _
	or inAddressEx(strAddress,"西村山郡.*大江町.*樽山") _
	or inAddressEx(strAddress,"西村山郡.*大江町.*沢口") _
	or inAddressEx(strAddress,"西村山郡.*大江町") then
		Yamagata = "山形県 西村山郡大江町（柳川、貫見、月布、十八才、樽山、沢口）" _
				 & "<br>冬期配達不可（11月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"西置賜郡.*飯豊町.*岩倉") _
	or inAddressEx(strAddress,"西置賜郡.*飯豊町.*宇津沢") _
	or inAddressEx(strAddress,"西置賜郡.*飯豊町.*上原") _
	or inAddressEx(strAddress,"西置賜郡.*飯豊町.*遅谷") _
	or inAddressEx(strAddress,"西置賜郡.*飯豊町.*数馬") _
	or inAddressEx(strAddress,"西置賜郡.*飯豊町.*上地谷") _
	or inAddressEx(strAddress,"西置賜郡.*飯豊町.*川内戸") _
	or inAddressEx(strAddress,"西置賜郡.*飯豊町.*白川") _
	or inAddressEx(strAddress,"西置賜郡.*飯豊町.*高造路") _
	or inAddressEx(strAddress,"西置賜郡.*飯豊町.*広河原") then
'	or inAddressEx(strAddress,"西置賜郡.*飯豊町.*萩生") then
		Yamagata = "山形県 西置賜郡飯豊町（岩倉、宇津沢、上原、遅谷、数馬、上地谷、川内戸、白川、高造路、広河原）" _
				 & "<br>冬期配達不可（11月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"東置賜郡.*川西町.*玉庭") then
'	or inAddressEx(strAddress,"東置賜郡.*高畠町") then
		Yamagata = "山形県 東置賜郡川西町大字玉庭6984（サンマリーナ玉庭）" _
				 & "<br>冬期配達不可（11月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"東置賜郡.*高畠町.*竹森") _
	or inAddressEx(strAddress,"東置賜郡.*高畠町.*二井宿") then
'	or inAddressEx(strAddress,"東置賜郡.*高畠町") then
		Yamagata = "山形県 東置賜郡高畠町（竹森、二井宿）" _
				 & "<br>冬期配達不可（11月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"東村山郡.*山辺町.*北山") _
	or inAddressEx(strAddress,"東村山郡.*山辺町.*大蕨") _
	or inAddressEx(strAddress,"東村山郡.*山辺町.*北作") _
	or inAddressEx(strAddress,"東村山郡.*山辺町.*畑谷") then
'	or inAddressEx(strAddress,"東村山郡.*山辺町.*大塚") then
		Yamagata = "山形県 東村山郡山辺町（大字北山、大字大蕨、大字北作、大字畑谷）" _
				 & "<br>冬期配達不可（11月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"尾花沢市.*鶴子") _
	or inAddressEx(strAddress,"尾花沢市.*銀山温泉") _
	or inAddressEx(strAddress,"尾花沢市.*行沢") _
	or inAddressEx(strAddress,"尾花沢市.*畑沢") _
	or inAddressEx(strAddress,"尾花沢市.*細野") _
	or inAddressEx(strAddress,"尾花沢市.*原田") _
	or inAddressEx(strAddress,"尾花沢市.*延沢") _
	or inAddressEx(strAddress,"尾花沢市.*鶴巻田") _
	or inAddressEx(strAddress,"尾花沢市.*舟生") _
	or inAddressEx(strAddress,"尾花沢市.*二藤袋") _
	or inAddressEx(strAddress,"尾花沢市.*六沢") _
	or inAddressEx(strAddress,"尾花沢市.*母袋") _
	or inAddressEx(strAddress,"尾花沢市.*正厳") _
	or inAddressEx(strAddress,"尾花沢市.*上柳渡戸") _
	or inAddressEx(strAddress,"尾花沢市.*下柳渡戸") _
	or inAddressEx(strAddress,"尾花沢市.*中島") _
	or inAddressEx(strAddress,"尾花沢市.*牛房野") _
	or inAddressEx(strAddress,"尾花沢市.*北郷") _
	or inAddressEx(strAddress,"尾花沢市.*押切") _
	or inAddressEx(strAddress,"尾花沢市.*毒沢") _
	or inAddressEx(strAddress,"尾花沢市.*高橋") _
	or inAddressEx(strAddress,"尾花沢市.*市野") _
	or inAddressEx(strAddress,"尾花沢市.*岩谷沢") _
	or inAddressEx(strAddress,"尾花沢市.*南沢") _
	or inAddressEx(strAddress,"尾花沢市.*富山") _
	or inAddressEx(strAddress,"尾花沢市.*寺内") _
	or inAddressEx(strAddress,"尾花沢市.*五十沢") then
'	or inAddressEx(strAddress,"尾花沢市.*横町") then
		Yamagata = "山形県 尾花沢市（鶴子・銀山温泉・行沢・畑沢・細野・原田・延沢・鶴巻田" _
				 & "・舟生・二藤袋・六沢・母袋・正厳・上柳渡戸・下柳渡戸・中島・牛房野" _
				 & "・北郷・押切・毒沢・高橋・市野々・岩谷沢・南沢・富山・寺内・五十沢444～579）" _
				 & "<br>冬期配達不可（12月1日～3月末）"
		exit function
	end if
	if inAddressEx(strAddress,"北村山郡.*大石田町.*次年子") _
	or inAddressEx(strAddress,"北村山郡.*大石田町.*大浦") _
	or inAddressEx(strAddress,"北村山郡.*大石田町.*駒籠") _
	or inAddressEx(strAddress,"北村山郡.*大石田町.*桂木町") _
	or inAddressEx(strAddress,"北村山郡.*大石田町.*豊田") _
	or inAddressEx(strAddress,"北村山郡.*大石田町.*横山") _
	or inAddressEx(strAddress,"北村山郡.*大石田町.*四日市") _
	or inAddressEx(strAddress,"北村山郡.*大石田町.*田沢") _
	or inAddressEx(strAddress,"北村山郡.*大石田町.*川前") then
'	or inAddressEx(strAddress,"村山郡.*朝日町") then
		Yamagata = "山形県 北村山郡大石田町（次年子・大浦・駒籠・桂木町・豊田・横山・四日市・田沢・川前）" _
				 & "<br>冬期配達不可（12月1日～3月末）"
		exit function
	end if
	if inAddressEx(strAddress,"最上郡.*最上町.*堺田") _
	or inAddressEx(strAddress,"最上郡.*最上町.*満澤") _
	or inAddressEx(strAddress,"最上郡.*最上町.*月楯") _
	or inAddressEx(strAddress,"最上郡.*最上町.*冨澤") _
	or inAddressEx(strAddress,"最上郡.*最上町.*東法田") _
	or inAddressEx(strAddress,"最上郡.*最上町.*向町") _
	or inAddressEx(strAddress,"最上郡.*最上町.*志茂") _
	or inAddressEx(strAddress,"最上郡.*最上町.*大堀") _
	or inAddressEx(strAddress,"最上郡.*大蔵町.*南山") _
	or inAddressEx(strAddress,"最上郡.*大蔵町.*赤松") _
	or inAddressEx(strAddress,"最上郡.*鮭川村.*曲川") _
	or inAddressEx(strAddress,"最上郡.*鮭川村.*大芦沢") _
	or inAddressEx(strAddress,"最上郡.*鮭川村.*羽根沢") _
	or inAddressEx(strAddress,"最上郡.*鮭川村.*中渡") _
	or inAddressEx(strAddress,"最上郡.*戸沢村.*角川") _
	or inAddressEx(strAddress,"最上郡.*戸沢村.*古口") _
	or inAddressEx(strAddress,"最上郡.*戸沢村.*蔵岡") _
	or inAddressEx(strAddress,"最上郡.*真室川町.*差首鍋") _
	or inAddressEx(strAddress,"最上郡.*真室川町.*及位") _
	or inAddressEx(strAddress,"最上郡.*真室川町.*釜淵") _
	or inAddressEx(strAddress,"最上郡.*真室川町.*川.内") _
	or inAddressEx(strAddress,"最上郡.*真室川町.*木.下") _
	or inAddressEx(strAddress,"最上郡.*真室川町.*大沢") _
	or inAddressEx(strAddress,"最上郡.*真室川町.*大滝") _
	or inAddressEx(strAddress,"最上郡.*金山町.*有屋") _
	or inAddressEx(strAddress,"最上郡.*金山町.*中田") _
	or inAddressEx(strAddress,"最上郡.*金山町.*安沢") _
	or inAddressEx(strAddress,"最上郡.*金山町.*飛森") _
	or inAddressEx(strAddress,"最上郡.*金山町.*大字下野明") _
	or inAddressEx(strAddress,"最上郡.*金山町.*山崎三枝") _
	or inAddressEx(strAddress,"最上郡.*金山町.*金山") _
	or inAddressEx(strAddress,"最上郡.*金山町.*上台") _
	or inAddressEx(strAddress,"最上郡.*金山町.*山崎") then
		Yamagata = "山形県 最上郡最上町（堺田・満澤・月楯・冨澤1～200番台・600番～1000番・富澤3780-1" _
				 & "・東法田・向町1500～2000番台・志茂900～1000番台・大堀1360-19（保養センターもがみ））" _
				 & "、大蔵町（南山・赤松）、鮭川村（曲川・大芦沢・羽根沢・中渡）" _
				 & "、戸沢村（角川、古口・大字蔵岡）、真室川町（差首鍋・及位・釜淵・川ノ内・木ノ下・大沢・大滝）" _
				 & "、金山町（有屋・中田・安沢・飛森・大字下野明・山崎三枝、金山900～2100、上台30～1100・山崎11～1600）" _
				 & "<br>冬期配達不可（12月1日～3月末）"
		exit function
	end if
	if inAddressEx(strAddress,"東根市.*猪野沢") _
	or inAddressEx(strAddress,"東根市.*泉郷") _
	or inAddressEx(strAddress,"東根市.*観音寺") _
	or inAddressEx(strAddress,"東根市.*関山") _
	or inAddressEx(strAddress,"東根市.*沼沢") then
'	or inAddressEx(strAddress,"東根市.*六田.楯.越") then
		Yamagata = "山形県 東根市（猪野沢、泉郷、観音寺、関山、沼沢）" _
				 & "<br>冬期配達不可（12月1日～3月31日）"
		exit function
	end if
	if inAddressEx(strAddress,"鶴岡市.*荒沢") _
	or inAddressEx(strAddress,"鶴岡市.*大網") _
	or inAddressEx(strAddress,"鶴岡市.*大鳥") _
	or inAddressEx(strAddress,"鶴岡市.*小名部") _
	or inAddressEx(strAddress,"鶴岡市.*上田沢") _
	or inAddressEx(strAddress,"鶴岡市.*木野俣") _
	or inAddressEx(strAddress,"鶴岡市.*倉沢") _
	or inAddressEx(strAddress,"鶴岡市.*小菅野代") _
	or inAddressEx(strAddress,"鶴岡市.*下田沢") _
	or inAddressEx(strAddress,"鶴岡市.*関川") _
	or inAddressEx(strAddress,"鶴岡市.*田麦俣") _
	or inAddressEx(strAddress,"鶴岡市.*羽黒町手向") _
	or inAddressEx(strAddress,"鶴岡市.*宝谷") _
	or inAddressEx(strAddress,"鶴岡市.*松沢") then
'	or inAddressEx(strAddress,"鶴岡市.*伊勢原町") then
		Yamagata = "山形県 鶴岡市（荒沢､大網､大鳥､小名部､上田沢､木野俣､倉沢､小菅野代､下田沢､関川､田麦俣､羽黒町手向､宝谷､松沢）" _
				 & "<br>冬期配達不可（12月1日～3月31日）"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'熊本県
'-----------------------------------------------------------------------
Private Function Kumamoto(byVal strAddress)
	Kumamoto = ""
	if inAddress(strAddress,"球磨郡")	then
	if inAddress(strAddress,"五木村")	then
		Kumamoto = "熊本県 球磨郡（五木村）<br>週２回(火、金）配達"
		exit function
	end if
	end if
	if inAddress(strAddress,"八代市")	then
	if inAddress(strAddress,"泉町")	then
		Kumamoto = "熊本県 八代市泉町（泉町下岳除く）<br>冬季配達不可（１２月～３月）"
		exit function
	end if
	end if
End Function
'-----------------------------------------------------------------------
'大分県
'-----------------------------------------------------------------------
Private Function Oita(byVal strAddress)
	Oita = ""
	if inAddress(strAddress,"大分市")	then
	if inAddress(strAddress,"中ノ洲、西ノ洲")	then
		Oita = "大分県 大分市（中ノ洲2番地、3番地－昭和電工業内各企業宛　西ノ洲1番地－新日本製鉄（株）内各企業宛）<br>重量物はチャータ扱い"
		exit function
	end if
	end if
	if inAddress(strAddress,"日田市")	then
	if inAddress(strAddress,"前津江町、中津江村、上津江町、大山町、天ヶ瀬町")	then
		Oita = "大分県 日田市（前津江町、中津江村、上津江町、大山町、天ヶ瀬町）<br>パレット、重厚長大チャーター"
		exit function
	end if
	end if
End Function

'-----------------------------------------------------------------------
'宮崎県
'-----------------------------------------------------------------------
Private Function Miyazaki(byVal strAddress)
	Miyazaki = ""

	if inAddress(strAddress,"宮崎市")	then
	if inAddress(strAddress,"鏡洲")	then
		Miyazaki = "宮崎県 宮崎市大字鏡洲 工事現場宛<br>支店止め又はチャーター"
		exit function
	end if
	end if

	if inAddress(strAddress,"西都市、児湯郡")	then
	if inAddress(strAddress,"上場、銀鏡、尾八重、八重、中尾、片内、西米良村")	then
		Miyazaki = "宮崎県 西都市上場、銀鏡、尾八重、八重、中尾、片内・児湯郡西米良村<br>週３回 (月･水・金)配達"
		exit function
	end if
	end if

	if inAddress(strAddress,"児湯郡")	then
	if inAddress(strAddress,"木城町")	then
	if inAddress(strAddress,"石河内・中之又・川原")	then
		Miyazaki = "宮崎県 児湯郡木城町石河内・中之又・川原<br>週３回 (火･木・土)配達"
		exit function
	end if
	end if
	end if

	if inAddress(strAddress,"児湯郡")	then
	if inAddress(strAddress,"木城町")	then
	if inAddress(strAddress,"石河内・中之又・川原")	then
		Miyazaki = "宮崎県 児湯郡木城町石河内・中之又・川原<br>週３回 (火･木・土)配達"
		exit function
	end if
	end if
	end if

	if inAddress(strAddress,"東臼杵郡")	then
	if inAddress(strAddress,"椎葉村")	then
	if inAddress(strAddress,"大河内")	then
'		Miyazaki = "宮崎県 東臼杵郡（椎葉村大河内、美郷町南郷区上渡川）<br>週２回（月、木）配達"
		Miyazaki = "宮崎県 東臼杵郡椎葉村大河内<br>週２回（月、木）配達"
		exit function
	end if
	end if
	end if

	if inAddress(strAddress,"東臼杵郡")	then
	if inAddress(strAddress,"椎葉村、諸塚村")	then
	if inAddress(strAddress,"松尾、下福良、七ツ山、立岩")	then
'		Miyazaki = "宮崎県 東臼杵郡椎葉村（松尾、下福良５７０～１６９９、諸塚村七ツ山・立岩・十根川・仲塔）<br>週２回（火、金）配達"
		Miyazaki = "宮崎県 東臼杵郡椎葉村松尾、椎葉村下福良５７０～１６９９、諸塚村七ツ山・立岩<br>週２回（火、金）配達"
		exit function
	end if
	end if
	end if

	if inAddress(strAddress,"東臼杵郡")	then
	if inAddress(strAddress,"椎葉村")	then
	if inAddress(strAddress,"不土野、下福良")	then
'		Miyazaki = "宮崎県 東臼杵郡（椎葉村不土野、椎葉村下福良　１９００～２５００、椎葉村尾前）<br>週２回（水、土）配達"
		Miyazaki = "宮崎県 東臼杵郡椎葉村不土野、椎葉村下福良　１９００～２５００<br>週２回（水、土）配達"
		exit function
	end if
	end if
	end if

End Function
'-----------------------------------------------------------------------
'沖縄県
'-----------------------------------------------------------------------
Private Function Okinawa(byVal strAddress)
	Okinawa = ""
	if inAddress(strAddress,"南城市")	then
	if inAddress(strAddress,"知念")	then
	if inAddress(strAddress,"久高")	then
		Okinawa = "沖縄県 南城市知念字久高<br>港止め（港からの配達不可）"
		exit function
	end if
	end if
	end if

	if inAddress(strAddress,"うるま市、宮古島市、南城市")	then
	if inAddress(strAddress,"津堅、伊良部町、知念字久高")	then
		Okinawa = "沖縄県 うるま市津堅、宮古島市伊良部町、南城市知念字久高<br>港止（配達不能）出港（月～土）"
		exit function
	end if
	end if

	if inAddress(strAddress,"島尻郡")	then
	if inAddress(strAddress,"北大東村、南大東村")	then
		Okinawa = "沖縄県 島尻郡北大東村、南大東村<br>不定期月５回出航"
		exit function
	end if
	end if

	if inAddress(strAddress,"島尻郡")	then
	if inAddress(strAddress,"粟国村、伊是名村、伊平屋村、座間味村、渡嘉敷村、渡名喜村")	then
		Okinawa = "沖縄県 島尻郡粟国村、伊是名村、伊平屋村、座間味村、渡嘉敷村、渡名喜村<br>港止（配達不能）出港（月～土）"
		exit function
	end if
	end if

	if inAddress(strAddress,"多良間村、八重山郡")	then
		Okinawa = "沖縄県 宮古郡多良間村、八重山郡全域<br>港止（配達不能）出港（週４回）"
		exit function
	end if

End Function
'-----------------------------------------------------------------------
'アドレスキーワードチェック 3
'-----------------------------------------------------------------------
Private Function inAddress3(byVal strAddress,byVal strKey1,byVal strKey2,byVal strKey3)
	inAddress3 = False
	dim	i
	for i = 1 to 3
		dim	strKey
		select case i
		case 1
			strKey = strKey1
		case 2
			strKey = strKey2
		case 3
			strKey = strKey3
		end select		
		if strKey = "" then
			exit for
		end if
		inAddress3 = inAddress(strAddress,strKey)
		if inAddress3 = False then
			exit for
		end if
	next
End Function
'-----------------------------------------------------------------------
'高知県
'-----------------------------------------------------------------------
Private Function Kochi(byVal strAddress)
	Kochi = ""
	if	inAddress(strAddress,"宿毛市沖の島町") then	',四万十市具同
		Kochi = "高知県 宿毛市沖の島町<br>配達不能地区"
		exit function
	end if
	if	inAddress(strAddress,"四万十町興津") then	',高岡郡越知町
		Kochi = "高知県 高岡郡四万十町興津<br>週２，３回の配達"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'愛媛県
'-----------------------------------------------------------------------
Private Function Ehime(byVal strAddress)
	Ehime = ""
	if inAddress(strAddress,"新居浜市別子山,四国中央市富郷町,四国中央市金砂町") then	',四国中央市妻鳥町
		Ehime = "愛媛県	新居浜市別子山・四国中央市富郷町○○、金砂町○○<br>週２回（月・木）配達"
		exit function
	end if
	if (inAddress(strAddress,"宇和島市津島町") > 0 and inAddress(strAddress,"曲烏、須下、成漁家、鰹網代") > 0) _
	or (inAddress(strAddress,"南宇和郡愛南町") > 0 and inAddress(strAddress,"網代、家串、魚神山、平碆、油袋") > 0) then
'	or (inAddress(strAddress,"宇和島市") > 0 and inAddress(strAddress,"大浦甲") > 0) then
		Ehime = "愛媛県 宇和島市津島町の内(曲烏、須下、成漁家、鰹網代)・南宇和郡愛南町網代、家串、魚神山、平碆、油袋<br>週２回（水・金）配達"
		exit function
	end if
	if	inAddress(strAddress,"西予市野村町惣川") then
'	or	inAddress(strAddress,"西予市宇和町卯之町") then
		Ehime = "愛媛県	西予市野村町惣川<br>週３回(月･水・金)配達"
		exit function
	end if
	if	inAddress(strAddress,"越智郡上島町生名") then	',越智産業
		Ehime = "愛媛県	越智郡上島町生名<br>支店止め又はチャーター"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'香川県
'-----------------------------------------------------------------------
Private Function Kagawa(byVal strAddress)
	Kagawa = ""
	if	inAddress(strAddress,"小豆郡土庄町豊島") then	',小豆郡土庄町
		Kagawa = "香川県 小豆郡土庄町豊島<br>実重量31kg以上の荷物は配達不能"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'徳島県
'-----------------------------------------------------------------------
Private Function Tokushima(byVal strAddress)
	Tokushima = ""
	if	inAddress(strAddress,"吉野川市美郷村") _
	or	inAddress(strAddress,"美馬市木屋平") _
	or	inAddress(strAddress,"三好市東祖谷、三好市西祖谷山村") _
	or	inAddress(strAddress,"美馬郡つるぎ町") then	'、美馬市美馬町
		Tokushima = "徳島県	吉野川市美郷村・美馬市木屋平・三好市東祖谷○○、西祖谷山村・美馬郡つるぎ町一宇<br>週３回(月･水・金)配達"
		exit function
	end if
	if inAddress(strAddress,"勝浦郡上勝町") then	',板野郡北島町鯛浜字西ノ須
		Tokushima = "徳島県	勝浦郡上勝町<br>週３回(月･水・金)配達※パレット配達不可"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'山口県
'-----------------------------------------------------------------------
Private Function Yamaguchi(byVal strAddress)
	Yamaguchi = ""
	if inAddress(strAddress,"下関市")			then
		if inAddress(strAddress,"蓋井島、六連島")	then
			Yamaguchi = "山口県 下関市蓋井島・六連島<br>配達不能地区"
			exit function
		end if
	end if
	if inAddress(strAddress,"周南市")	then
		if inAddress(strAddress,"金峰・須万・須金・中須北")	then
			Yamaguchi = "山口県 周南市（金峰・須万・須金・中須北）<br>"
			Yamaguchi = Yamaguchi + "週1～2回配達（不定期）"
			Yamaguchi = Yamaguchi + "※軽四車両での配達の為、1件10ｋｇ未満、1個30ｋｇ未満のみ配達可能。長尺物・パレットは配達不可。"
			exit function
		end if
	end if
	if inAddress(strAddress,"岩国市")	then
		if inAddress(strAddress,"三角町")	then
			if inAddress(strAddress,"岩国米軍基地")	then
				Yamaguchi = "山口県 岩国市三角町○○丁目　岩国米軍基地　ＥＸゾーン宛・オンベースエリアの現場宛<br>営業所止め"
				exit function
			end if
		end if
	end if
	if inAddress(strAddress,"岩国市")	then
		if inAddress(strAddress,"愛宕山")	then
			Yamaguchi = "山口県 岩国市愛宕山地内　岩国飛行場　愛宕山低層住宅現場宛<br>営業所止め"
		end if
	end if
End Function
'-----------------------------------------------------------------------
'広島県
'-----------------------------------------------------------------------
Private Function Hiroshima(byVal strAddress)
	Hiroshima = ""
	if inAddress(strAddress,"広島") = False then
		exit function
	end if
	if inAddress(strAddress,"尾道市向東町")	then	'、尾道市美ノ郷町
		Hiroshima = "広島県 尾道市向東町　加島宛<br>支店止"
		exit function
	end if
	if inAddress(strAddress,"尾道市因島重井町")	then	'、尾道市美ノ郷町
		Hiroshima = "広島県 尾道市因島重井町　小細島宛<br>支店止"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'鳥取県
'-----------------------------------------------------------------------
Private Function Tottori(byVal strAddress)
	Tottori = ""
	if inAddress(strAddress,"鳥取") = False then
		exit function
	end if
	if inAddress(strAddress,"八頭郡若桜町つく米") then
'	or inAddress(strAddress,"八頭郡八頭町橋本")	then
		Tottori = "鳥取県 八頭郡若桜町つく米<br>冬季配達不能（11月～4月）"
		exit function
	end if
	if inAddress(strAddress,"西伯郡")	then	',八頭郡
		if inAddress(strAddress,"伯耆町")	then
			if inAddress(strAddress,"岩立、桝水、金屋谷")	then
				Tottori = "鳥取県 西伯郡伯耆町（岩立、桝水、金屋谷）<br>冬季配達不能（12月～3月）"
			end if
		elseif inAddress(strAddress,"大山町")	then	',八頭町
			if inAddress(strAddress,"大山、豊房、香取")	then	',橋本
				Tottori = "鳥取県 西伯郡大山町（大山、豊房、香取）<br>冬季配達不能（12月～3月）"
			end if
		end if
		exit function
	end if
	if inAddress(strAddress,"日野郡")	then
	if inAddress(strAddress,"江府町")	then
	if inAddress(strAddress,"御机、鏡ヶ成")	then
		Tottori = "鳥取県 日野郡江府町（御机、鏡ヶ成）<br>冬季配達不能（12月～3月）"
		exit function
	end if
	end if
	end if
	if inAddress(strAddress,"日野郡")	then	',八頭郡
		if inAddress(strAddress,"日野郡江府町御机,日野郡江府町鏡")	then
			Tottori = "鳥取県 日野郡江府町（御机、鏡ヶ成）<br>冬季配達不能（12月～3月）"
		elseif inAddress(strAddress,"日野郡日野町久住")	then
			Tottori = "鳥取県 日野郡日野町久住<br>冬季配達不能（12月～3月）"
		elseif inAddress(strAddress,"日野郡日南町神戸上")	then	',八頭郡八頭町橋本
			Tottori = "鳥取県 日野郡日南町神戸上<br>冬季配達不能（12月～3月）"
		end if
		exit function
	end if
	if inAddress(strAddress,"境港市竹内団地") then	',境港市昭和町
		if inAddress(strAddress,"６１、６２")	then	'、１１
			Tottori = "鳥取県 境港市竹内団地61・62<br>支店止め又はチャーター"
		else
			Tottori = "鳥取県 境港市竹内団地71<br>支店止め又はチャーター"
		end if
		exit function
	end if
	if inAddress(strAddress,"米子市流通町")	then	',米子市蚊屋
		Tottori = "鳥取県 米子市流通町430-25 国分西日本㈱米子総合センター<br>支店止め又はチャーター"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'島根県
'-----------------------------------------------------------------------
Private Function Shimane(byVal strAddress)
	Shimane = ""
	if inAddress(strAddress,"宍道町佐々布")	then	',宍道町西来
		Shimane = "島根県 宍道町佐々布939-2　えびす本郷㈱松江営業所<br>支店止め又はチャーター"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'岡山県
'-----------------------------------------------------------------------
Private Function Okayama(byVal strAddress)
	Okayama = ""
	if inAddress(strAddress,"岡山") = False then
		exit function
	end if
	if inAddress(strAddress,"北区大内田")	then
		if inAddress(strAddress,"６７５、コンベックス")	then	'、８３０
			Okayama = "岡山県 岡山市北区大内田675　コンベックス岡山<br>支店止め又はチャーター"
		end if
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'奈良県
'-----------------------------------------------------------------------
Private Function Nara(byVal strAddress)
	Nara = ""
	if inAddress(strAddress,"奈良") = False then
		exit function
	end if
	if inAddress(strAddress,"野迫川村")	then	'、櫟本町
		Nara = "奈良県 吉野郡野迫川村<br>週１～２回配達"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'兵庫県
'-----------------------------------------------------------------------
Private Function Hyogo(byVal strAddress)
	Hyogo = ""
	if inAddress(strAddress,"兵庫") = False then
		exit function
	end if
	if inAddress(strAddress,"神戸市中央区港島８")	then	'、神戸市中央区港島１
		Hyogo = "兵庫県	神戸市中央区港島8-7　㈱築港　神戸事業所　及び　㈱築港　ポートアイランド化学品センター<br>支店止め又はチャーター"
		exit function
	end if
	if inAddress(strAddress,"神戸市灘区")	then
		if inAddress(strAddress,"大月台、摩耶山町、六甲山町")	then	'、大石南町
			Hyogo = "兵庫県	神戸市灘区(大月台、摩耶山町、六甲山町）<br>週３回（月・水・金）配達"
		end if
		if inAddress(strAddress,"摩耶埠頭")	then	'
			Hyogo = "兵庫県	神戸市灘区摩耶埠頭2-1　五洋港運株式会社神戸物流センター及び五洋海運株式会社<br>支店止め又はチャーター"
		end if
		exit function
	end if
	if inAddress(strAddress,"養父市")	then
		if inAddress(strAddress,"氷ノ山、大久保、丹戸、福定、別宮、奈良尾")	then	'、中瀬
			Hyogo = "兵庫県	養父市鉢伏（ハチ高原）、氷ノ山スキー場、大久保、丹戸、福定、別宮、奈良尾<br>冬季配達不能（12月～3月）"
		end if
		exit function
	end if
	if inAddress(strAddress,"村岡区大笹、小代区新屋") then	'、揖保郡太子町
		Hyogo = "兵庫県	美方郡香美町村岡区（丸味・大笹）、小代区新屋<br>冬季配達不能（12月～3月）"
		exit function
	end if
	if inAddress(strAddress,"新温泉町春来") then	'、揖保郡太子町
		Hyogo = "兵庫県	美方郡新温泉町春来<br>冬季配達不能（12月～3月）"
		exit function
	end if
	if inAddress(strAddress,"神河町上小田") then	'、市川町西川辺
		Hyogo = "兵庫県	神崎郡神河町上小田881-146 峰山高原<br>冬季配達不能（12月～3月）"
		exit function
	end if
	if inAddress(strAddress,"南あわじ市沼島") then	'、南あわじ市山添
		Hyogo = "兵庫県	南あわじ市沼島<br>配達不可（営業所止め）"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'大阪府
'-----------------------------------------------------------------------
Private Function Osaka(byVal strAddress)
	Osaka = ""
	if inAddress(strAddress,"住之江区")			then
	if inAddress(strAddress,"インテックス大阪")	then
		Osaka = "大阪府	大阪市住之江区…インテックス大阪会場宛てのみ<br>配達不能地区"
		exit function
	end if
	end if
	if inAddress(strAddress,"大阪市")	then
	if inAddress(strAddress,"西区")		then
	if inAddress(strAddress,"ドーム")	then
	if inAddress(strAddress,"スタジアムモール")	then
		Osaka = "大阪府	大阪市西区千代崎3-2-1 大阪ドーム内（スタジアムモール宛は配達可）<br>支店止め又はチャーター"
		exit function
	end if
	end if
	end if
	end if
End Function
'-----------------------------------------------------------------------
'和歌山県
'-----------------------------------------------------------------------
Private Function Wakayama(byVal strAddress)
	Wakayama = ""
	if inAddress(strAddress,"和歌山") = False then
		exit function
	end if
	if inAddress(strAddress,"伊都郡") then	'・大谷
		if (inAddress(strAddress,"かつらぎ町")  and inAddress(strAddress,"花園")) _
		or (inAddress(strAddress,"高野町") and inAddress(strAddress,"西富貴・上筒香・下筒香")) then
			Wakayama = "和歌山県 伊都郡かつらぎ町（花園○○）、高野町（東冨貴・西富貴・上筒香・中筒香・下筒香）<br>週３回(火・木・土)配達"
		end if
		exit function
	end if
	if inAddress(strAddress,"那智勝浦町") then
		if inAddress(strAddress,"大野、樫原、口色川、熊瀬川、小阪、小匠、坂足、田垣内、那智山、南大居")	then	'、朝日
			Wakayama = "和歌山県 東牟婁郡那智勝浦町（大野､樫原､口色川､熊瀬川､小阪､小匠､坂足､田垣内､那智山､南大居）<br>週２回配達"
		end if
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'京都府
'-----------------------------------------------------------------------
Private Function Kyoto(byVal strAddress)
	Kyoto = ""
	if inAddress(strAddress,"京都") = False then
		exit function
	end if
	if inAddress(strAddress,"右京区") then
		if inAddress(strAddress,"嵯峨越畑、嵯峨水尾、嵯峨樒原") then	'、花園
			Kyoto = "京都府 右京区（嵯峨越畑○○・嵯峨水尾○○・嵯峨樒原○○）<br>配達不能地区"
		end if
		exit function
	end if
	if inAddress(strAddress,"北区")	then
		if inAddress(strAddress,"大森、杉坂、雲ヶ畑、真弓")	then	'、西賀茂
			Kyoto = "京都府 北区（大森・杉坂・雲ヶ畑・真弓）<br>配達不能地区"
			exit function
		end if
	end if
	if inAddress(strAddress,"西京区大原野出灰町,西京区大原野外畑町") then	',西京区大原野上里北ノ町
		Kyoto = "京都府 西京区（大原野出灰町,大原野外畑町）<br>配達不能地区"
		exit function
	end if
	if inAddress(strAddress,"左京区") then
		if inAddress(strAddress,"大原、鞍馬、久多、花背、広河原、静市静原、大原百井町、北白川地蔵谷町")	then	'、岡崎徳成町
			Kyoto = "京都府 左京区大原○○,鞍馬○○,久多○○,花背○○,広河原○○,静市静原,大原百井町、北白川地蔵谷町<br>配達不能地区"
		end if
		exit function
	end if
	if inAddress(strAddress,"福知山市") then
		if inAddress(strAddress,"天座・石場・一尾・市寺・一ノ宮・行積・猪野々・今安・梅谷・漆端・榎原・夷・大江町・大呂・奥野部・上天津・上大内・上小田・上佐々木・上野条・鴨野町・喜多・北山・雲原・瘤木・小牧・下天津・下大内・下小田・下野条・十二・常願寺・大門・樽水・田和・談・長尾・中佐々木・野花・拝師・筈巻・畑中・半田・日尾・牧・宮垣・室・夜久野町・和久寺") then	'、岡崎徳成町
			Kyoto = "京都府 福知山市（天座・石場・一尾・市寺・一ノ宮・行積・猪野々・今安・梅谷・漆端・榎原・夷・大江町・大呂・奥野部・上天津・上大内・上小田・上佐々木・上野条・鴨野町・喜多・北山・雲原・瘤木・小牧・下天津・下大内・下小田・下野条・十二・常願寺・大門・樽水・田和・談・長尾・中佐々木・野花・拝師・筈巻・畑中・半田・日尾・牧・宮垣・室・夜久野町・和久寺）" _
				  & "<br>週1～２回配達"
		end if
		exit function
	end if
	if inAddress(strAddress,"伏見区稲荷山")	then
		Kyoto = "京都府伏見区稲荷山官有地、稲荷山○○<br>週1～２回配達"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'三重県
'-----------------------------------------------------------------------
Private Function Mie(byVal strAddress)
	Mie = ""
	if inAddress(strAddress,"三重") = False then
		exit function
	end if
	if inAddress(strAddress,"津市榊原町") then
'	or inAddress(strAddress,"津市高茶屋小森上野町") then
		Mie = "三重県 津市榊原町4183-12　航空自衛隊笠取山分屯基地<br>週１回配達"
		exit function
	end if
	if inAddress(strAddress,"四日市市山之一色") then
'	or inAddress(strAddress,"四日市市日永西４") then
		if inAddress(strAddress,"加藤産業") then
			Mie = "三重県 四日市市山之一色町800-5　加藤産業株式会社<br>支店止め又はチャーター"
			exit function
		end if
	end if
	if inAddress(strAddress,"四日市市中村町") _
	or inAddress(strAddress,"四日市市山之一色") then
'	or inAddress(strAddress,"四日市市堀木") then
		if inAddress(strAddress,"東芝") then	' 、草川工業
			Mie = "三重県 四日市市中村町2549-3、四日市市山之一色800" _
				& "㈱東芝四日市工場及び敷地内工事現場<br>支店止め又はチャーター"
			exit function
		end if
	end if
	if inAddress(strAddress,"川越町亀先新田") then
'	or inAddress(strAddress,"菰野町大字永井") then
		if inAddress(strAddress,"前田運送") then	'、中堀電商
			Mie = "三重県 三重郡川越町亀先新田120　前田運送株式会社<br>支店止め又はチャーター"
			exit function
		end if
	end if
	if inAddress(strAddress,"川越町高松") then
'	or inAddress(strAddress,"菰野町大字永井") then
		if inAddress(strAddress,"東海国分、山星屋、トライアル") then	'、中堀電商
			Mie = "三重県 三重郡川越町高松1515-1　東海国分㈱、㈱山星屋、トライアル三重川越常温センター<br>支店止め又はチャーター"
			exit function
		end if
	end if
	if inAddress(strAddress,"桑名市大字福岡町") _
	or inAddress(strAddress,"桑名市福岡町") then
'	or inAddress(strAddress,"桑名市大字江場") then
		if inAddress(strAddress,"菱食") then	'、毛利きね
			Mie = "三重県 桑名市大字福岡町475-1　株式会社菱食桑名物流センター<br>支店止め又はチャーター"
			exit function
		end if
	end if
	if inAddress(strAddress,"志摩市磯部町渡鹿野") _
	or inAddress(strAddress,"志摩市磯部町渡的矢") _
	or inAddress(strAddress,"志摩市磯部町三") then
'	or inAddress(strAddress,"志摩市阿児町") then
		Mie = "三重県 志摩市磯部町渡鹿野、磯部町的矢、磯部町三ヶ所<br>配達不能地区"
	end if
	if inAddress(strAddress,"紀宝町浅里") _
	or inAddress(strAddress,"紀宝町神内") _
	or inAddress(strAddress,"紀宝町高岡") _
	or inAddress(strAddress,"紀宝町平尾井") _
	or inAddress(strAddress,"紀宝町桐原") _
	or inAddress(strAddress,"紀宝町阪松原") then
'	or inAddress(strAddress,"御浜町大字阿田和") then
		Mie = "三重県 南牟婁郡紀宝町…浅里・神内・高岡・平尾井・桐原・阪松原の各地区<br>配達不能地区"
	end if
End Function
'-----------------------------------------------------------------------
'滋賀県
'-----------------------------------------------------------------------
Private Function Shiga(byVal strAddress)
	Shiga = ""
	if inAddress(strAddress,"滋賀") = False then
		exit function
	end if
	if inAddress(strAddress,"近江八幡市沖島") then
'	or inAddress(strAddress,"近江八幡市末広町") then
		Shiga = "滋賀県	近江八幡市沖島町<br>配達不能地区"
		exit function
	end if
	if inAddress(strAddress,"大津市坂本本町") then
'	or inAddress(strAddress,"大津市仰木") then
		if inAddress(strAddress,"４２２０") then
			Shiga = "滋賀県	大津市坂本本町4220<br>支店止め又はチャーター"
		end if
		if inAddress(strAddress,"４２４４") then
			Shiga = "滋賀県	大津市坂本本町4244<br>支店止め又はチャーター"
		end if
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'愛知県
'-----------------------------------------------------------------------
Private Function Aichi(byVal strAddress)
	Aichi = ""
	if inAddress(strAddress,"愛知") = False then
		exit function
	end if
	if inAddress(strAddress,"名古屋市東区大幸南１")	then
'	or inAddress(strAddress,"名古屋市名東区文教台２")	then
		Aichi = "愛知県 名古屋市東区大幸南1丁目1-1　ナゴヤドーム各テナント宛<br>支店止め又はチャーター"
		exit function
	end if
	if inAddress(strAddress,"名古屋市名東区姫若町４０")	then
'	or inAddress(strAddress,"名古屋市名東区文教台２")	then
		Aichi = "愛知県 名古屋市名東区姫若町40" _
			  & "　佐川急便名古屋ＳＲＣ３階" _
			  & "　㈱井田両国堂名古屋支店佐川急便名古屋ＳＲＣ５階" _
			  & "　㈱サンワテクノス　名古屋店" _
			  & "<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"名古屋市千種区吹上２")	then
'	or inAddress(strAddress,"名古屋市千種区星が丘")	then
		Aichi = "愛知県 名古屋市千種区吹上2-6-3名古屋中小企業振興会館、吹上ホール各テナント宛<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"名古屋市熱田区熱田西町") then
'	or inAddress(strAddress,"名古屋市熱田区中田町")	then
		Aichi = "愛知県 名古屋市熱田区熱田西町1-1　名古屋国際会議場<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"名古屋市西区牛島町") then
'	or inAddress(strAddress,"名古屋市西区笠取町") then
		Aichi = "愛知県 名古屋市西区牛島町6-1　名古屋ルーセントタワー<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"春日井市明知町") then
'	or inAddress(strAddress,"春日井市小野町") then
		Aichi = "愛知県 春日井市明知町1514-88" _
			  & "オークワ東海食品センター、" _
			  & "加藤産業春日井センター、" _
			  & "国分中部オークワ東海食品センター、" _
			  & "モリモト東海資材センター、" _
			  & "サンライズ東海物流センター、" _
			  & "パーティハウス、" _
			  & "三菱食品オークワ東海食品センター、" _
			  & "山星屋オークワ東海食品センター" _
			  & "<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"東海市東海町") then
'	or inAddress(strAddress,"東海市名和町") then
		Aichi = "愛知県 東海市東海町5-3" _
			  & "　日本製鉄㈱名古屋製鐵所及び敷地内工事現場" _
			  & "<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"東海市浅山") then
'	or inAddress(strAddress,"東海市名和町") then
		Aichi = "愛知県 東海市浅山2-47" _
			  & "　アスクル名古屋センター" _
			  & "<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"常滑市セントレア") then
'	or inAddress(strAddress,"常滑市大鳥町") then
		Aichi = "愛知県 常滑市セントレア5-10-1" _
			  & "　愛知県国際展示場" _
			  & "<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"稲沢市石橋") then
'	or inAddress(strAddress,"稲沢市一色下方町") then
		Aichi = "愛知県 稲沢市石橋4-1-1" _
			  & "　三菱食品株式会社" _
			  & "<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"知多市南浜町") then
'	or inAddress(strAddress,"知多市八幡") then
		Aichi = "愛知県 知多市南浜町11" _
			  & "　出光興産㈱愛知製油所及び敷地内工事現場" _
			  & "<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"北名古屋市沖村権現") then
'	or inAddress(strAddress,"北名古屋市九之坪北浦") then
		Aichi = "愛知県 北名古屋市沖村権現35-1" _
			  & "、大宝運輸㈱西春支店内、国分中部㈱西春センター" _
			  & "<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"金城埠頭") then
'	or inAddress(strAddress,"名古屋市天白区") then
		Aichi = "愛知県 金城埠頭２丁目２" _
			  & " 名古屋国際展示場「ポートメッセ名古屋」" _
			  & "<br>チャータ扱い"
	end if
	if inAddress(strAddress,"みよし市打越町山") then
'	or inAddress(strAddress,"みよし市三好町中鯰") then
		Aichi = "愛知県 みよし市打越町山ノ神10-1" _
			  & "<br>○○ドミーみよしセンター・国分㈱三河流通センター" _
			  & "<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"一宮市今伊勢町馬寄川田") then
'	or inAddress(strAddress,"一宮市本町４") then
		Aichi = "愛知県 一宮市今伊勢町馬寄川田16" _
			  & "、バロー一宮トーカンセンター" _
			  & "<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"一宮市千秋町天摩") then
'	or inAddress(strAddress,"一宮市赤見３") then
		Aichi = "愛知県 一宮市千秋町天摩字金島1" _
			  & "　赤ちゃん本舗愛知ＩＤＣ、伊藤忠食品㈱一宮東物流センター" _
			  & "<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"一宮市奥町風田") then
'	or inAddress(strAddress,"一宮市赤見３") then
		Aichi = "愛知県 一宮市奥町風田27-1・奥町風田40-1" _
			  & "　シモハナ物流株式会社一宮第２センター" _
			  & "<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"長久手市茨") then
'	or inAddress(strAddress,"長久手市下山") then
		Aichi = "愛知県 長久手市茨ヶ廻間乙1533-1" _
			  & "　愛・地球博記念公園（モリコロパーク）　イベント会場・出展ブース宛" _
			  & "<br>支店止め又はチャーター"
	end if
	if inAddress(strAddress,"豊橋市") then
	if inAddress(strAddress,"中部採石、三嶽鉱山、鶴田石材、東海アスコン、照山砕石、アスレック、東海シーエス、東海パーツ") then ' 、神野新田町
		Aichi = "愛知県 豊橋市（中部採石、三嶽鉱山、鶴田石材、東海アスコン、照山砕石、アスレック、東海シーエス、東海パーツ）" _
			  & "<br>配達不能地区"
	end if
	end if
End Function
'-----------------------------------------------------------------------
'千葉県
'-----------------------------------------------------------------------
Private Function Chiba(byVal strAddress)
	Chiba = ""
	if inAddressEx(strAddress,"美浜区.*中瀬.*幕張メッセ") then
		Chiba = "千葉県 美浜区中瀬２－１幕張ﾒｯｾ日本ｺﾝﾍﾞﾝｼｮﾝｾﾝﾀｰ<br>各ﾃﾅﾝﾄ宛チャータ扱い"
		exit function
	end if
	if inAddressEx(strAddress,"美浜区.*浜田.*海浜幕張") then
'	or inAddressEx(strAddress,"美浜区.*中瀬.*９.*１") then
		Chiba = "千葉県 美浜区浜田2-102 海浜幕張パーキングエリア内<br>支店止め又はチャーター"
		exit function
	end if
	if inAddressEx(strAddress,"中央区.*浜野町.*１０２５") then
'	or inAddressEx(strAddress,"中央区.*都町.*１８") then
		Chiba = "千葉県 中央区浜野町1025-150 日酒販アマゾン京葉流通センター<br>支店止め又はチャーター"
		exit function
	end if
	if inAddressEx(strAddress,"船橋市.*浜町.*３") then
'	or inAddressEx(strAddress,"船橋市.*薬円台.*５") then
		Chiba = "千葉県 船橋市浜町3丁目全域<br>支店止め又はチャーター"
		exit function
	end if
	if inAddressEx(strAddress,"芝山町.*香山新田") then
'	or inAddressEx(strAddress,"芝山") then
		Chiba = "千葉県 芝山町香山新田93-4 新東京国際空港 工事現場宛<br>支店止め又はチャーター"
		exit function
	end if
	if inAddressEx(strAddress,"芝山町.*香山新田.*雨堤") then
'	or inAddressEx(strAddress,"芝山") then
		Chiba = "千葉県 芝山町香山新田字雨堤76 成田国際空港振興協会<br>支店止め又はチャーター"
		exit function
	end if
	if inAddressEx(strAddress,"袖.*浦市.*中袖.*１.*１") then
'	or inAddressEx(strAddress,"袖.*浦市.*野里") then
		Chiba = "千葉県 袖ヶ浦市中袖１－１ 東京ガス構内<br>支店止め又はチャーター"
		exit function
	end if
	if inAddressEx(strAddress,"袖.*浦市.*代宿") _
	or inAddressEx(strAddress,"袖.*浦市.*椎.*森") then
'	or inAddressEx(strAddress,"袖.*浦市.*野里") then
		Chiba = "千葉県 袖ヶ浦市代宿・椎の森・椎の森工業団地内の工事現場<br>支店止め又はチャーター"
		exit function
	end if
	if inAddressEx(strAddress,"市原市.*青柳") then
'	or inAddressEx(strAddress,"市原市.*能満") then
		Chiba = "千葉県 市原市青柳１～2100番地<br>※51ｋｇ以上は、支店止め又はチャーター"
		exit function
	end if
	if inAddressEx(strAddress,"木更津市.*中島.*地先") then
'	or inAddressEx(strAddress,"木更津市.*中島") then
		Chiba = "千葉県 木更津市中島地先　海ほたる（アクアライン海ほたるパーキング）<br>週４回(月･水･金･土）配達"
		exit function
	end if
	if inAddressEx(strAddress,"木更津市.*築地.*１") then
		Chiba = "千葉県 木更津市築地1番地" _
					& " 新日鐵住金（株）君津製鐵所本館 君津製鉄所ビジネスセンター 東洋スチレン（株）新日化学木更津製鐵所各企業宛" _
					& "<br>支店止め又はチャーター"
		exit function
	end if
	if inAddressEx(strAddress,"君津市.*君津.*１")	then
		Chiba = "千葉県 君津市君津1番地 新日鐵住金（株）君津製鐵所<br>支店止め又はチャーター"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'神奈川県
'-----------------------------------------------------------------------
Private Function Kanagawa(byVal strAddress)
	Kanagawa = ""
	if inAddress(strAddress,"横浜市")	then
	if inAddress(strAddress,"西区")		then
	if inAddress(strAddress,"みなとみらい")	then
	if inAddress(strAddress,"パシフィコ横浜")	then
		Kanagawa = "神奈川県 横浜市西区みなとみらい１－１－１ ﾊﾟｼﾌｨｺ横浜（展示ﾎｰﾙ・ｱﾈｯｸｽﾎｰﾙ・会議ｾﾝﾀｰ・国立大ﾎｰﾙ）<br>チャータ扱い"
		exit function
	end if
	end if
	end if
	end if
	if inAddress(strAddress,"横浜市")	then
	if inAddress(strAddress,"中区")		then
	if inAddress(strAddress,"本牧ふ頭１")	then
	if inAddress(strAddress,"横浜港運事業協同組合")	then
		Kanagawa = "神奈川県 横浜市中区本牧ふ頭1番地 横浜港運事業協同組合<br>支店止"
		exit function
	end if
	end if
	end if
	end if
	if inAddress(strAddress,"横浜市")	then
	if inAddress(strAddress,"鶴見区")		then
	if inAddress(strAddress,"扇島")	then
		Kanagawa = "神奈川県 横浜市鶴見区扇島2-1・3-1・4-1 JFE スチール㈱東日本製鉄所内 東京ガス㈱扇島工場<br>配達不能地区"
		exit function
	end if
	end if
	end if
	if inAddress(strAddress,"横浜市")	then
	if inAddress(strAddress,"鶴見区")		then
	if inAddress(strAddress,"大黒町")	then
	if inAddress(strAddress,"築港横浜化学品")	then
		Kanagawa = "神奈川県 横浜市鶴見区大黒町5-81、大黒町9-15　㈱築港横浜化学品センター第１倉庫、第２倉庫<br>支店止め又はチャーター"
		exit function
	end if
	end if
	end if
	end if
	if inAddress(strAddress,"川崎市")	then
	if inAddress(strAddress,"東扇島")	then
	if inAddress(strAddress,"２３")	then
	if inAddress(strAddress,"１０")	then
		Kanagawa = "神奈川県 川崎市東扇島23-10 ワールドサプライ第一センター内　三菱食品東扇島百貨店物流センター、国分首都圏㈱、及び㈱池利<br>支店止め又はチャーター"
		exit function
	end if
	end if
	end if
	end if
	if inAddress(strAddress,"伊勢原市")	then
	if inAddress(strAddress,"大山")	then
		Kanagawa = "神奈川県 伊勢原市大山<br>チャータ扱い"
		exit function
	end if
	end if
	if inAddress(strAddress,"海老名市")	then
	if inAddress(strAddress,"東柏ヶ谷")	then
	if inAddress(strAddress,"厚木航空施設司令部")	then
		Kanagawa = "神奈川県 海老名市東柏ヶ谷 在日米海軍厚木航空施設司令部内<br>支店止め又はチャーター"
		exit function
	end if
	end if
	end if
	if inAddress(strAddress,"座間市")	then
	if inAddress(strAddress,"陸軍基地管理本部")	then
		Kanagawa = "神奈川県 座間市座間 在日米陸軍基地管理本部 キャンプ座間内<br>支店止め又はチャーター"
		exit function
	end if
	end if
	if inAddress(strAddress,"在日米海軍厚木航空施設司令部")	then
		Kanagawa = "神奈川県 上草柳、下草柳、福田、本蓼川 在日米海軍厚木航空施設司令部内<br>支店止め又はチャーター"
		exit function
	end if
	'TEST
'	if inAddress(strAddress,"川崎市")	then
'	if inAddress(strAddress,"麻生区")	then
'		Kanagawa = "神奈川県<br>TEST"
'		exit function
'	end if
'	end if
End Function
'-----------------------------------------------------------------------
'東京都
'-----------------------------------------------------------------------
Private Function Tokyo(byVal strAddress)
	Tokyo = ""
	if inAddress(strAddress,"綾瀬市") > 0 and inAddress(strAddress,"深谷、蓼川、本蓼川、大上") > 0 then
		Tokyo = "東京都 綾瀬市（深谷、蓼川、本蓼川、大上）在日米海軍厚木航空施設司令部内<br>支店止め又はチャーター"
		exit function
	end if
	if inAddress(strAddress,"足柄上郡山北町") > 0 and inAddress(strAddress,"皆瀬川・川西") > 0 then
		Tokyo = "東京都 足柄上郡山北町皆瀬川・川西900番地～<br>支店止め又はチャーター"
		exit function
	end if
	if inAddress(strAddress,"八丈支庁") > 0 then
		Tokyo = "東京都 八丈支庁…………鳥島<br>配達不能地区"
		exit function
	end if
	if inAddress(strAddress,"江東区有明３") > 0 then
		Tokyo = "東京都 江東区有明3-11-1東京国際展示場、3-10-1東京ﾋﾞｯｸﾞｻｲﾄ内各ﾃﾅﾝﾄ宛<br>チャータ扱い"
		exit function
	end if
	if inAddress(strAddress,"江東区青海") > 0  and inAddress(strAddress,"東京ビッグサイト") > 0 then
		Tokyo = "東京都 江東区青海1-2-33　東京ビッグサイト青海展示棟<br>支店止め又はチャーター"
		exit function
	end if
	if inAddress(strAddress,"江東区青海２") > 0 _
	or(inAddress(strAddress,"江東区青海３") > 0 and inAddress(strAddress,"中央防波堤") > 0)_
	or inAddress(strAddress,"江東区豊洲２") > 0 _
	or inAddress(strAddress,"江東区豊洲６") > 0 _
	or inAddress(strAddress,"江東区有明１") > 0 _
	or inAddress(strAddress,"江東区有明２") > 0 then
		Tokyo = "東京都 江東区青海2丁目防波堤(地先)、3丁目中央防波堤（地先）、豊洲2全域（現場宛のみ）"
		Tokyo = Tokyo & "豊洲6全域（現場宛のみ）、有明1全域（現場宛のみ）、有明2全域（現場宛のみ）<br>チャータ扱い"
		exit function
	end if
	if inAddress(strAddress,"世田谷区") > 0 then
		Tokyo = "東京都 世田谷区宛の重量物（1件100kg以上）のみ<br>支店止め又はチャーター"
		exit function
	end if
	if inAddress(strAddress,"千代田区丸の内") > 0  and inAddress(strAddress,"大丸東京") > 0 then
		Tokyo = "東京都 千代田区丸の内1-9-1　大丸東京店<br>配達不能地区"
		exit function
	end if
	if inAddress(strAddress,"中央区豊海") > 0  and inAddress(strAddress,"冷蔵庫") > 0 then
		Tokyo = "東京都 中央区豊海　○○冷蔵庫宛<br>チャータ扱い"
		exit function
	end if
	if inAddress(strAddress,"大田区") > 0  and inAddress(strAddress,"羽田空港") > 0 then
		Tokyo = "東京都 大田区羽田空港宛の重量物（1件100kg以上）のみ<br>支店止め又はチャーター"
		exit function
	end if
	if inAddressEx(strAddress,"大田区城南島.*５.*１.*１") > 0 then
		Tokyo = "東京都 大田区城南島5-1-1　TOCP内兼松新東亜食品㈱<br>支店止め又はチャーター"
		exit function
	end if
	if inAddress(strAddress,"大田区平和島６") > 0 then
		Tokyo = "東京都 大田区平和島（6-2-1、6-2-25、6-3-1、6-3-2）  東京団地冷蔵株式会社<br>支店止め又はチャーター"
		exit function
	end if
	if inAddress(strAddress,"昭島市拝島町") > 0 then
		Tokyo = "東京都 昭島市拝島町3927-7（いなげや国分昭島センター・ＩＫＤ国分昭島センター）<br>支店止め又はチャーター"
		exit function
	end if
	if inAddressEx(strAddress,"昭島市武蔵野.*物流") > 0 then
		Tokyo = "東京都 昭島市武蔵野2-9-18（株式会社オリンピック昭島物流・株式会社キララ昭島物流）<br>支店止め又はチャーター"
		exit function
	end if
	if inAddress(strAddress,"八王子市石川町２９６９") > 0 then
		Tokyo = "東京都 八王子市石川町2969-18ヒューテックノーリン東京支店内　関東惣菜RDC<br>チャータ扱い"
		exit function
	end if
	if inAddress(strAddress,"立川市緑町") > 0 then
		Tokyo = "東京都 立川市緑町　国営昭和記念公園内各露店宛て<br>チャータ扱い"
		exit function
	end if
	if inAddress(strAddress,"西多摩郡") > 0 and inAddress(strAddress,"奥多摩町，桧原村") > 0 then
		Tokyo = "東京都 西多摩郡奥多摩町，桧原村<br>週１回配達　300ｋｇ以上チャータ扱い"
		exit function
	end if
	if inAddressEx(strAddress,"西多摩郡瑞穂町.*二本木.*４６１") > 0 then
		Tokyo = "東京都 西多摩郡瑞穂町二本木461-2"
		Tokyo = Tokyo & "イトーヨーカドー西多摩共配センター、"
		Tokyo = Tokyo & "IY西多摩加食共配センター、"
		Tokyo = Tokyo & "三井食品IY西多摩加食共配センター、"
		Tokyo = Tokyo & "㈱高山IY西多摩加食共配センター、"
		Tokyo = Tokyo & "日本酒類販売㈱IY西多摩加食センター、"
		Tokyo = Tokyo & "コンフェックスIY西多摩加食共配センター、"
		Tokyo = Tokyo & "あらたIY西多摩加食共配センター、"
		Tokyo = Tokyo & "ＩＹ西多摩ＩＤＣ"
		Tokyo = Tokyo & "<br>支店止め又はチャーター"
		exit function
	end if
	if inAddress(strAddress,"北区豊島５") > 0 then
		Tokyo = "東京都 北区豊島5-1-21　日本出版販売㈱　王子流通センター宛<br>支店止め又はチャーター"
		exit function
	end if
	if inAddress(strAddress,"品川区八潮") > 0 then
		Tokyo = "東京都 "
		Tokyo = Tokyo & "品川区八潮2-8-1　㈱宇徳"
		Tokyo = Tokyo & "品川区八潮2-9　㈱宇徳"
		Tokyo = Tokyo & "品川区八潮2-9　ジャパンエキスプレス"
		Tokyo = Tokyo & "品川区八潮2-6-2　日本通運㈱"
		Tokyo = Tokyo & "品川区八潮2-6-4　ケイヒン"
		Tokyo = Tokyo & "品川区八潮2-1-2　ダイトーコーポレーション"
		Tokyo = Tokyo & "<br>支店止め又はチャーター"
		exit function
	end if
	if inAddress(strAddress,"港区高輪３") > 0 then
		Tokyo = "東京都 "
		Tokyo = Tokyo & "港区高輪3-13-1 "
		Tokyo = Tokyo & "グランドプリンスホテル高輪、グランドプリンスホテル新高輪、ザ・プリンスさくらタワー東京"
		Tokyo = Tokyo & "グランドプリンスホテル国際パミール"
		Tokyo = Tokyo & "<br>支店止め又はチャーター"
		exit function
	end if
	if inAddress(strAddress,"港区高輪４") > 0 then
		Tokyo = "東京都 "
		Tokyo = Tokyo & "港区高輪4-10-30 品川プリンスホテル"
		Tokyo = Tokyo & "<br>支店止め又はチャーター"
		exit function
	end if
	if inAddress(strAddress,"町田市中町１") > 0 then
		Tokyo = "東京都 "
		Tokyo = Tokyo & "町田市中町1-20-23　シバヒロ　※シバヒロの記載が無い場合もあり"
		Tokyo = Tokyo & "<br>支店止め又はチャーター"
		exit function
	end if
End Function
'-----------------------------------------------------------------------
'福島県
'-----------------------------------------------------------------------
Private Function Fukushima(byVal strAddress)
	Fukushima = ""
	if inAddress(strAddress,"福島市") then
	if inAddress(strAddress,"土湯温泉町") then
	if inAddress(strAddress,"野地、鷲倉山、幕川、大笹、在庭坂、中ノ堂、金堀沢、町庭坂、高湯、湯花沢、神ノ森、砥石山、目洗川、蓬平") then
		Fukushima = "福島県 福島市土湯温泉町（野地、鷲倉山、幕川、大笹）、在庭坂（中ノ堂、金堀沢）、町庭坂（高湯、湯花沢、神ノ森、砥石山、目洗川、蓬平）<br>冬期配達不可（12月1日～3月31日）"
	end if
	end if
	end if

'	if inAddress(strAddress,"福島市") then
'	if inAddress(strAddress,"桜本木通沢") then
'	if inAddress(strAddress,"信夫温泉") then
'	if inAddress(strAddress,"のんびり") then
'		Fukushima = "福島県 福島市桜本木通沢４　信夫温泉のんびり館<br>冬期配達不可（12月21日～3月31日）"
'	end if
'	end if
'	end if
'	end if

	if inAddress(strAddress,"いわき市") then
	if inAddress(strAddress,"藤原町") then
	if inAddress(strAddress,"蕨平") then
	if inAddress(strAddress,"５０") then
		Fukushima = "福島県 いわき市藤原町蕨平50　スパリゾートハワイアンズ<br>支店止め又はチャーター"
	end if
	end if
	end if
	end if
	if inAddress(strAddress,"南相馬市") then
	if inAddress(strAddress,"小高区") then
		Fukushima = "福島県 南相馬市小高区○○<br>支店止め又はチャーター"
	end if
	end if
	if inAddress(strAddress,"相馬市") then
	if inAddress(strAddress,"山上、玉野、東玉野") then
		Fukushima = "福島県 相馬市（山上、玉野、東玉野）<br>支店止め又はチャーター"
	end if
	end if
	if inAddress(strAddress,"相馬市") then
	if inAddress(strAddress,"光陽") then
	if inAddress(strAddress,"相馬エネルギーパーク") then
		Fukushima = "福島県 相馬市光陽2-2-21　相馬エネルギーパーク<br>支店止め又はチャーター"
	end if
	end if
	end if

	if inAddress(strAddress,"郡山市") then
	if inAddress(strAddress,"湖南町") then
		Fukushima = "福島県 郡山市湖南町〇〇<br>※冬期配達不能（12月～3月）"
	end if
	end if

	if inAddress(strAddress,"伊達郡") then
	if inAddress(strAddress,"川俣町") then
	if inAddress(strAddress,"山木屋") then
		Fukushima = "福島県 伊達郡川俣町（山木屋○○）<br>冬期配達不可（12月21日～3月末）"
	end if
	end if
	end if
	if inAddress(strAddress,"相馬郡") then
	if inAddress(strAddress,"新地町") then
	if inAddress(strAddress,"駒ケ岳今神") then
	if inAddress(strAddress,"新地火力発電所") then
		Fukushima = "福島県 相馬郡新地町駒ケ岳今神1-1　新地火力発電所<br>支店止め又はチャーター"
	end if
	end if
	end if
	end if
	if inAddress(strAddress,"相馬郡") then
	if inAddress(strAddress,"新地町") then
	if inAddress(strAddress,"駒ケ岳") then
	if inAddress(strAddress,"今神") then
	if inAddress(strAddress,"１５９－１") then
		Fukushima = "福島県 相馬郡新地町駒ケ岳今神159-1　ＬＮＧ基地（相馬港天然ガス発電所）<br>支店止め又はチャーター"
	end if
	end if
	end if
	end if
	end if
	if inAddress(strAddress,"相馬郡") then
	if inAddress(strAddress,"飯舘村") then
		Fukushima = "福島県 相馬郡飯舘村<br>支店止め又はチャーター"
	end if
	end if
	if inAddress(strAddress,"双葉郡") then
	if inAddress(strAddress,"大熊町、双葉町") then
		Fukushima = "福島県 双葉郡（大熊町、双葉町）<br>警戒区域の為、配達不能"
	end if
	end if
	if inAddress(strAddress,"双葉郡") then
	if inAddress(strAddress,"富岡町、浪江町、楢葉町、葛尾村、川内村") then
		Fukushima = "福島県 双葉郡（富岡町、浪江町、楢葉町、葛尾村、川内村）<br>支店止め又はチャーター"
	end if
	end if
	if inAddress(strAddress,"岩瀬郡") then
	if inAddress(strAddress,"天栄村") then
	if inAddress(strAddress,"田良尾、羽鳥、湯本") then
		Fukushima = "福島県 岩瀬郡天栄村（田良尾、羽鳥、湯本）<br>冬期配達不能(12月1日～3月31日)"
	end if
	end if
	end if

	if inAddress(strAddress,"田村市") then
	if inAddress(strAddress,"都路町") then
	if inAddress(strAddress,"古道") then
		Fukushima = "福島県 田村市都路町古道○○<br>週１回配達※冬期配達不能（12月～3月）"
	else
		Fukushima = "福島県 田村市都路町〇〇<br>※冬期配達不能（12月～3月）"
	end if
	end if
	end if

	if inAddress(strAddress,"南会津郡") then
	if inAddress(strAddress,"南会津町") then
	if inAddress(strAddress,"岩下数間沢") then
	if inAddress(strAddress,"３") then
		Fukushima = "福島県 南会津郡南会津町岩下数間沢3　住友金属鉱山㈱八総鉱山<br>支店止め又はチャーター"
	end if
	end if
	end if
	end if
End Function
'-----------------------------------------------------------------------
'宮城県
'2019.09.30
'-----------------------------------------------------------------------
Private Function Miyagi(byVal strAddress)
	Miyagi = ""
	if inAddress(strAddress,"仙台市") then
	if inAddress(strAddress,"泉区") then
	if inAddress(strAddress,"福岡") then
	if inAddress(strAddress,"岳山") then
		Miyagi = "宮城県 仙台市泉区福岡字岳山（泉高原スプリングバレースキー場、泉ヶ岳スキー場等関連施設)<br>冬期配達不可（12月1日～3月31日）"
		exit function
	end if
	end if
	end if
	end if
	if inAddress(strAddress,"気仙沼市") then
	if inAddress(strAddress,"浅根､磯草､浦の浜､大初平､大向､亀山､駒形､外畑､外浜､高井､田尻､長崎､中山､廻館､三作浜､要害､横沼") then
		Miyagi = "宮城県 気仙沼市（浅根､磯草､浦の浜､大初平､大向､亀山､駒形､外畑､外浜､高井､田尻､長崎､中山､廻館､三作浜､要害､横沼）<br>港船上渡し（前日に荷受人様に連絡が取れた荷物を船会社へ引き渡し）着払/代引不可"
		exit function
	end if
	end if
	if inAddress(strAddress,"大崎市") then
	if inAddress(strAddress,"鳴子温泉鬼首") then
		Miyagi = "宮城県 大崎市鳴子温泉鬼首<br>冬期配達不可（12月1日～3月31日）"
		exit function
	end if
	if inAddress(strAddress,"栗原市") then
	if inAddress(strAddress,"栗駒沼倉耕英") then
		Miyagi = "宮城県 栗原市栗駒沼倉耕英〇〇<br>支店止め又はチャーター"
		exit function
	end if
	end if
	if inAddress(strAddress,"大崎市") then
	if inAddress(strAddress,"鳴子温泉鬼首地熱発電所,地熱発電所") then
		Miyagi = "宮城県 大崎市鳴子温泉鬼首地熱発電所<br>配達不能地区"
		exit function
	end if
	end if
	end if
	if inAddress(strAddress,"刈田郡") then
	if inAddress(strAddress,"七ヶ宿町") then
		Miyagi = "宮城県 刈田郡七ヶ宿町<br>週２～３回配達"
		exit function
	end if
	end if
End Function
'-----------------------------------------------------------------------
'埼玉県
'-----------------------------------------------------------------------
Private Function Saitama(byVal strAddress)
	Saitama = ""
	if inAddress(strAddress,"秩父市") then
	if inAddress(strAddress,"大滝、中津川、三峰") then
		Saitama = "埼玉県 秩父市大滝、中津川、三峰<br>週2～３回配達"
		exit function
	end if
	end if
	if inAddress(strAddress,"熊谷市") then
	if inAddress(strAddress,"千代") then
	if inAddress(strAddress,"ヤオコー、国分関信越、加藤産業、日本アクセス、日酒販、伊藤忠食品、三菱食品") then
		Saitama = "埼玉県 熊谷市千代703-1・株式会社ヤオコー・国分関信越株式会社・加藤産業株式会社・株式会社日本アクセス・日酒販株式会社・伊藤忠食品株式会社・三菱食品株式会社<br>支店止め又はチャーター"
		exit function
	end if
	end if
	end if
	if inAddress(strAddress,"所沢市") then
	if inAddress(strAddress,"牛沼西保戸窪") then
	if inAddress(strAddress,"キューソー流通センター") then
		Saitama = "埼玉県 所沢市牛沼西保戸窪489-3のキューソー流通センター内各企業<br>支店止め又はチャーター"
	end if
	end if
	end if
	if inAddress(strAddress,"飯能市") then
	if inAddress(strAddress,"赤沢、吾野、井上、大河原、上赤工、上直竹上分、上直竹下分、上長沢、上名栗、上畑、唐竹、苅生、北川、久須美、小岩井、虎秀、小瀬戸、坂石、坂石町分、坂元、下赤江、下直竹、下名栗、下畑、白子、高山、長沢、永田、永田台、中藤上郷、中藤中郷、中藤下郷、原市場、平戸、南、南川") then
		Saitama = "埼玉県 飯能市（赤沢・吾野・井上・大河原・上赤工・上直竹上分・上直竹下分・上長沢・上名栗・上畑・唐竹・苅生・北川・久須美・小岩井・虎秀・小瀬戸・坂石・坂石町分・坂元・下赤江・下直竹・下名栗・下畑・白子・高山・長沢・永田・永田台・中藤上郷・中藤中郷・中藤下郷・原市場・平戸・南・南川）<br>週1回配達（火曜日）"
	end if
	end if
	if inAddress(strAddress,"朝霞市") then
	if inAddress(strAddress,"上内間木") then
	if inAddress(strAddress,"４５９－１") then
		Saitama = "埼玉県 朝霞市上内間木459-1のＢＬＳ朝霞センター<br>支店止め又はチャーター"
	end if
	end if
	end if
	if inAddress(strAddress,"入間市") then
	if inAddress(strAddress,"宮寺宮ノ台") then
	if inAddress(strAddress,"４１０２-３５") then
		Saitama = "埼玉県 入間市宮寺宮ノ台4102-35　㈱ロジスティクス・ネットワーク、㈱若菜ロジネット、㈱東京ニチレイサービス<br>支店止め又はチャーター"
	end if
	end if
	end if
	if inAddress(strAddress,"川越市") then
	if inAddress(strAddress,"下赤坂") then
	if inAddress(strAddress,"５９３－１") then
		Saitama = "埼玉県 川越市下赤坂593-1　（株）CGCジャパングロサリ広域センター<br>支店止め（着店チャーター不可）"
	end if
	end if
	end if
	if inAddress(strAddress,"川越市") then
	if inAddress(strAddress,"下赤坂") then
	if inAddress(strAddress,"１８２２－１") then
		Saitama = "埼玉県 川越市下赤坂1822-1　カインズ川越センター<br>支店止め又はチャーター"
	end if
	end if
	end if
	if inAddress(strAddress,"狭山市") then
	if inAddress(strAddress,"根岸宇田木前") then
	if inAddress(strAddress,"６７７－１") then
		Saitama = "埼玉県 狭山市根岸宇田木前677-1（株式会社ヤオコー・国分株式会社・加藤産業株式会社・三菱食品株式会社）<br>支店止め又はチャーター"
	end if
	end if
	end if
End Function
Private Function inAddress(byVal strAddress,byVal strKeyword)
	inAddress = false
	dim	strDlm
	strDlm = "、"
	dim	strKey
	if inStr(strKeyword,"・") then
		strDlm = "・"
	end if
	if inStr(strKeyword,",") then
		strDlm = ","
	end if
	for each strKey in Split(strKeyword,strDlm)
		inAddress = inStr(1,strAddress,strKey,vbTextCompare)
		if inAddress then
			exit function
		end if
	next
End Function
'--------------------------------------------------------------
'http://www.tohoho-web.com/js/regexp.htm
'--------------------------------------------------------------
Private Function inAddressEx(byVal strAddress,byVal strKeyword)
	dim	regEx
	set	regEx = New RegExp
	regEx.Pattern = strKeyword
	regEx.Global = False
	regEx.IgnoreCase = True
	inAddressEx = regEx.Test(strAddress)
	set	regEx = Nothing
End Function
Private Function Han2Zen(byVal strSrc)
	dim	objB
	set objB = Server.CreateObject("Basp21")
	Han2Zen = objB.HAN2ZEN(strSrc)
	Han2Zen = strSrc
	set objB = Nothing
End Function
Private Function inStr8003(byVal strAddress)
	inStr8003 = false
	dim	lngPos
	lngPos = inStr(strAddress,"8")
	if lngPos = 0 then
		lngPos = inStr(strAddress,"８")
	end if
	if lngPos > 0 then
		dim	i
		for i = 1 to 3
			if isNumeric(Mid(strAddress,lngPos + i,1)) = false then
				exit function
			end if
		next
		if isNumeric(Mid(strAddress,lngPos + i,1)) = false then
			inStr8003 = true
		end if
		exit function
	end if
End Function
Private Function inStr1000(byVal strAddress)
	inStr1000 = false
	dim	lngPos
	lngPos = inStr(strAddress,"1")
	if lngPos = 0 then
		lngPos = inStr(strAddress,"１")
	end if
	if lngPos > 0 then
		dim	i
		for i = 1 to 3
			if isNumeric(Mid(strAddress,lngPos + i,1)) = false then
				exit function
			end if
		next
		if isNumeric(Mid(strAddress,lngPos + i,1)) = false then
			inStr1000 = true
		end if
		exit function
	end if
End Function
%>
