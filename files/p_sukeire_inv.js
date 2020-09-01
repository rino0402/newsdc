/*
p_sukeire_inv.js
2017.11.16 商品化納品リスト
\\w4\newsdc\exe\sei0010
*/
//var	version = '2017.11.16 ベータ版';
//var	version = '2017.11.30 箱代：済／工料：済／外装：対応中／加工費：未';
//var	version = '2017.12.01 正式対応版';
//var	version = '2017.12.05 課題対応中<br><small>・１回目の検索が10分以上かかる(未)<br>・行がずれる場合がある(未)<br>・塗り潰しが印刷されない(対応中)</small>';
var	vie11 = '<p>Internet Explorer 11 でヘッダーを固定すると表示がズレます。<br>ヘッダーを固定する場合は他のブラウザ(Chrome,Firefox)を使用して下さい。</p>';
var	arr = [];
arr.push('<p>2020.07.20 v.017<br>'
		+'・冷蔵庫請求書 BU加工費 引取除外<br>');
arr.push('<p>2020.06.16 v.016a<br>'
		+'・冷蔵庫請求書 BU加工費請求書<br>'
		+'・仕向先：01 がセットされる不具合修正');
arr.push('<p>2019.09.25 v.015<br>'
		+'・合計数量追加');
arr.push('<p>2019.03.27 v.014<br>'
		+'・印刷時、選択行の背景を無効<br>'
		+'・HE（給湯品番）の背景色を白黒印刷で判別しやすい黄色に変更');
arr.push('<p>2018.06.12 v.013<br>'
		+'・並び順変更<br>'
		+'　変更前：見積差額の大きい順<br>'
		+'　変更後：受入日、品番 順');
arr.push('<p>2018.06.01 v.012<br>'
		+'・検索速度Up(※若干です^^;)<br>'
		+'・画面レイアウト変更(※検索ボタンを右側に変更)<br>'
		+'・数字カンマ区切り<br>'
		+'・見積必要分を先頭に表示※差額の大きい順)');
arr.push('<p>2018.04.12 v.011</p>'
		+'・合計行追加');
arr.push('<p>2018.04.04 v.010</p>'
		+'・収単/担当者で検索できない不具合を修正');
arr.push('<p>2018.03.28 v.009</p>'
		+'・Copyボタンで結果をExcelに貼付できるようにしました.');
arr.push('<p>2018.03.08 v.008</p>'
		+'・受入数量=0 を検索しないように修正');
arr.push('<p>2018.02.23 v.007</p>'
		+'・冷蔵庫 引取対応:備考欄に「引*取*」が入っている品番を数値として集計');
arr.push('<p>2018.02.22 v.006</p>'
		+'・冷蔵庫 引取対応:備考欄に「引取*」が入っている品番を数値として集計<br>'
		+'・CGIをPythonに変更');
arr.push('<p>2018.02.20 v.005</p>'
		+'・冷蔵庫 BU加工費対応<br>'
		+'・検索条件(仕向先、収単/担当者、見積チェック)を記憶');
arr.push('<p>2018.01.18</p>'
		+'・検索速度Up(改善中)1日分 50秒→20秒に改善<br>'
		+'・行がずれる(再現待ち)');
arr.push('<p>2018.01.16</p>'
		+'・検索速度Up(改善中)');
arr.push('<p>2018.01.15</p>'
		+'・検索速度Up→(暫定)処理中に経過時間を表示<br>'
		+'・テーブルヘッダー固定');
arr.push('<p>2018.01.12</p>'
		+'・検索速度Up→(暫定)前回処理時間を表示<br>'
		+'・塗り潰しが印刷されるように修正<br>'
		+'・見積チェックをプルダウンに変更<br>'
		+'・見積チェックで要見積だけを検索する条件を追加<br>'
		+'・塗り潰しの色を見やすく変更※表示レイアウトを全般的に変更');

var	version = arr[0];
var	tHistory = arr.join('');

function log(t) {
	if(!t) {
		$('#log').html('');
	} else {
		$('#log').html($('#log').html() + t + '<br>');
	}
}
function Log(v) {
	console.log(v);
//	var	html = $('#log').html();
//	html += v + '<br>';
//	$('#log').html(html);
}
var nFormat = function(number,n) {
	var	result = '';
	if (number != 0) {
		result = parseFloat(number).toFixed(n);
//		result = Number(result).toLocaleString("ja-JP", {maximumSignificantDigits : 2});
		result = (new Intl.NumberFormat('ja-JP', { minimumFractionDigits: n }).format(result))
	}
    return result;
};
//設定：読込セット
/*
function setConfig(id,def) {
	console.log( 'setConfig(1):' + id + ',' + def );
	var v = localStorage.getItem(id);
	$('#log').html($('#log').html() + 'setConfig(2):' + id + ',v=' + v + ',def=' + def + '<br>');
	if(v == null) {
		$('#log').html($('#log').html() + 'set def<br>');
		v = def;
	}
	console.log( 'setConfig(3):' + id + ',' + v );
	$(id).val(v);
	return v;
}
*/
$(document).ready(function() {
	$('#key').text(location.pathname.replace(/\/[^/]*$/, '') + '/');
	$('a[href^="#"]').click(function() {
		$("html, body").animate({scrollTop:0}, 550, "swing");
	});
	var	startTm = new Date(Date.parse(storage('start_time')));
	var	endTm = new Date(Date.parse(storage('end_time')));
	$('#start_time').text(dateTimeFormat(startTm,true));
	$('#end_time').text(dateTimeFormat(endTm,true));
	$('#lastTm').text(parseInt((endTm - startTm)/1000) + '秒');
	var	lastTm = ' 処理時間: ' + $('#lastTm').text() + ' 前回(' + $('#lastTm').text() + ')';
	// バージョン
	$('#version').html(vie11 + tHistory);
	$('#msg').html("条件を指定して[検索]をクリックして下さい。 " + lastTm + '<br>' + version);
	// 初期値セット
	//設定：初期値
	$("#dns").blur(function() {
//		if($(this).val() == '') {
//			$(this).val(location.pathname.split('/')[1]);
//		}
//		$('#dns_span').text(setConfig('#' + this.id, $(this).val()));
		$('#dns_span').text($(this).val());
		var	key = $('#key').text() + '#cname';
		$('#cname').text(localStorage.getItem(key));

		var	req = 'jgyobu.py?dns=' + $(this).val() + '&jgyobu=0';
		log(req);
		fetch(req).then((res) => {
			return res.json();
		}).then((json) => {
			if(json.list.length > 0) {
				$('#cname').text(json.list[0].Name);
				localStorage.setItem(key ,$('#cname').text());
			} else {
				$('#cname').text('');
			}
		}).catch(function(err) {
			$('#cname').text(err);
		});
	});
	$('input[type="text"].config').each(function() {
		var	key = $('#key').text() + '#' + this.id;
		var	val = localStorage.getItem(key);
		log('getItem():' + key + ':' + val);
		if(!val) {
			switch(this.id) {
			case 'limit'		:	val = 300;	break;
			case 'dns'			:	val = location.pathname.split('/')[1];	break;
			case 'SHIMUKE_CODE'	:	val = '';	break;
			case 'S_TANTO'		:	val = '';	break;
			case 'CHECK'		:	val = '1';	break;
			}
		}
		$(this).val(val);
	});
	$(document).on("change",'input[type="text"].config', function() {
		var	key = $('#key').text() + '#' + this.id;
		log('setItem():' + key + ':' + $(this).val());
		localStorage.setItem(key ,$(this).val());
	});
	$('#dns').trigger('blur');
//	setConfig('#limit', 300);

	var now = new Date();
	var dt2 = now.getFullYear()+
		( "0"+( now.getMonth()+1 ) ).slice(-2)+
		( "0"+now.getDate() ).slice(-2);
	if ( now.getDate() < 21 ) {
		now.setMonth(now.getMonth() - 1);
	}
	var dt1 = now.getFullYear()+( "0"+( now.getMonth()+1 ) ).slice(-2)+	"21";
	$('#UKEIRE_DT1').val(dt1);
	$('#UKEIRE_DT2').val(dt2);
	//設定：初期値
//	setConfig('#SHIMUKE_CODE', '01');
//	setConfig('#S_TANTO', '');
//	setConfig('#CHECK', '1');
	// 検索ボタン
	var $table = $('#table');
    $('#List').click(function () {
//		localStorage.setItem('#SHIMUKE_CODE',$('#SHIMUKE_CODE').val());
//		localStorage.setItem('#S_TANTO',$('#S_TANTO').val());
//		localStorage.setItem('#CHECK',$('#CHECK').val());
		$('#log').html('click:' + this.id + '<br>');
		$('#msg').html('検索中<span class="gif-load">...</span>前回: ' + $('#lastTm').text());
//		$('#msg').addClass('gif-load');
		$('#tbody').find("tr").remove();
//		$('#table > tfoot').find("tr").remove();
		$table.floatThead('destroy');
		//経過時間
		var	startTm = new Date();
		var	elapsedTime = setInterval(function(){
			var	curTm = new Date();
			$('#msg').html('検索中<span class="gif-load">...</span>前回:' + $('#lastTm').text() + ' 経過:' + parseInt((curTm - startTm)/1000) + '秒');
		},1000);
		//検索条件
//		var	req = 'p_sukeire_inv.py?';
		var	req = 'mcheck.py?list=p_sukeire&dns=' + $('#dns').val();
		if($('#UKEIRE_DT1').val() != '') {
//			req += '&UKEIRE_DT1=' + $('#UKEIRE_DT1').val();
			req += '&syuka_st=' + $('#UKEIRE_DT1').val();
		}
		if($('#UKEIRE_DT2').val() != '') {
//			req += ('&UKEIRE_DT2=' + $('#UKEIRE_DT2').val());
			req += '&syuka_ed=' + $('#UKEIRE_DT2').val();
		}
		if($('#SHIMUKE_CODE').val() != '') {
			req += ('&SHIMUKE_CODE=' + $('#SHIMUKE_CODE').val());
		}
		if($('#S_TANTO').val() != '') {
			req += ('&S_TANTO=' + $('#S_TANTO').val());
		}
		if($('#CHECK').val() != '') {
			req += ('&CHECK=' + $('#CHECK').val());
		}
//			$('#msg').text('fetch():' + req);
		$('#log').html($('#log').html() + 'fetch():' + req + '<br>');
		fetch(req)
		.then( function ( res ) {
			//経過時間クリア
			clearInterval(elapsedTime);
			//fetch結果check
			var contentType = res.headers.get("content-type");
			console.log('contentType:' + contentType);
			$('#log').html($('#log').html() + 'contentType:' + contentType + '<br>');
			var	endTm = new Date();
			Log(startTm);
			Log(endTm);
			Log(parseInt((endTm - startTm)/1000) + '秒');
			$('#lastTm').text(parseInt((endTm - startTm)/1000) + '秒');
			$('#msg').text('処理時間: ' + $('#lastTm').text() );
			clearInterval(elapsedTime);
			storage('start_time',startTm);
			storage('end_time',endTm);
			if(contentType && contentType.indexOf("application/json") !== -1) {
				console.log('return json');
				return res.json();
//				} else if(contentType && contentType.indexOf("html/text") !== -1) {
//					return res.text();
			} else {
				$('#log').html($('#log').html() + '<br>' + res.text());
				$('#msg').text('fetch()error:' + contentType);
			}
		})
//			.then( function ( text ) {
//				$('#msg').text('fetch()text:' + text);
//			})
		.then( function ( json ) {
//				$('#msg').text('fetch()json:' + json);
			$('#log').html($('#log').html() + 'HTTP_HOST:' + json.HTTP_HOST + '<br>');
			$('#log').html($('#log').html() + 'REQUEST_URI:' + json.REQUEST_URI + '<br>');
			$('#log').html($('#log').html() + 'dns:' + json.dns + '<br>');
			$('#log').html($('#log').html() + 'UKEIRE_DT1:' + json.UKEIRE_DT1 + '<br>');
			$('#log').html($('#log').html() + 'UKEIRE_DT2:' + json.UKEIRE_DT2 + '<br>');
			$('#log').html($('#log').html() + 'SHIMUKE_CODE:' + json.SHIMUKE_CODE + '<br>');
			$('#log').html($('#log').html() + 'S_TANTO:' + json.S_TANTO + '<br>');
			$('#log').html($('#log').html() + 'CHECK:' + json.CHECK + '<br>');
			$('#log').html($('#log').html() + 'sql:' + json.sql + '<br>');
			$('#log').html($('#log').html() + 'error:' + json.error + '<br>');
			$('#log').html($('#log').html() + 'json.data.length:' + json.data.length + '<br>');
			var	tr = '';
			for ( var i = 0 ; i < json.data.length ; i++ ) {
				tr += '<tr>';
				tr += '<td class="number">' + (i + 1) + '</td>';
				//受入日
				tr += '<td class="date">' + json.data[i].UKEIRE_DT + '</td>';
				//収単/担当者
				tr += '<td class="date">' + json.data[i].S_TANTO + '</td>';
				//品番
				var cls = '';
				switch(json.data[i].L_KISHU1.slice(0,2)) {
				case 'CS':
				case 'CU':	break;
				default:	if(json.data[i].JGYOBU == 'A') {
								cls += ' bg-warning';
							}
							break;
				}
				tr += '<td class="' + cls +'" title="' + json.data[i].L_KISHU1 + '">' + json.data[i].KEY_HIN_NO + '</td>';
				//品名
				tr += '<td class="clip small">' + $('<span/>').text( json.data[i].HIN_NAME ).html() + '</td>';
				//受入数
				tr += '<td class="number">' + nFormat(json.data[i].Qty,0) + '</td>';
				//工料
				cls = 'number';
				txt = nFormat(json.data[i].PrcKoryo,2);
				if ($('#CHECK').val() != "0" && typeof json.data[i].NewKoryo != "undefined") {
					if(json.data[i].NewKoryo != json.data[i].PrcKoryo) {
						cls += ' bg-danger';
						txt += '<br>' + nFormat(json.data[i].NewKoryo,2);
					}
				}
				tr += '<td class="' + cls +'">' + txt + '</td>';
				//工料金額
				tr += '<td class="number">' + nFormat(json.data[i].Koryo,2) + '</td>';
				//箱代
				cls = 'number';
				txt = nFormat(json.data[i].PrcHako,2);
				if ($('#CHECK').val() != "0" && typeof json.data[i].NewHako != "undefined") {
					if(json.data[i].NewHako != json.data[i].PrcHako) {
						cls += ' bg-danger';
						txt += '<br>' + nFormat(json.data[i].NewHako,2);
					}
				}
				tr += '<td class="hakogai ' + cls +'">' + txt + '</td>';
				//箱代金額
				tr += '<td class="hakogai number">' + nFormat(json.data[i].Hako,2) + '</td>';
				//外装
				cls = 'number';
				txt = nFormat(json.data[i].PrcGaiso,2);
				if ($('#CHECK').val() != "0" && typeof json.data[i].NewGaiso != "undefined") {
					if(json.data[i].NewGaiso != json.data[i].PrcGaiso) {
						cls += ' bg-danger';
						txt += '<br>' + nFormat(json.data[i].NewGaiso,2);
					}
				}
				tr += '<td class="hakogai ' + cls +'">' + txt + '</td>';
				//外装金額
				tr += '<td class="hakogai number">' + nFormat(json.data[i].Gaiso,2) + '</td>';
				//加工
				cls = 'number';
				txt = nFormat(json.data[i].PrcKako,2);
				if ($('#CHECK').val() != "0" && typeof json.data[i].NewKako != "undefined") {
					if(json.data[i].NewKako != json.data[i].PrcKako) {
						cls += ' bg-danger';
						txt += '<br>' + nFormat(json.data[i].NewKako,2);
					}
				}
				tr += '<td class="' + cls +'">' + txt + '</td>';
				//加工金額
				tr += '<td class="number">' + nFormat(json.data[i].Kako,2) + '</td>';
				//BU加工
				cls = 'number';
				txt = nFormat(json.data[i].PrcBu,2);
				if ($('#CHECK').val() != "0" && typeof json.data[i].NewBu != "undefined") {
					if(json.data[i].NewBu != json.data[i].PrcBu) {
						cls += ' bg-danger';
						txt += '<br>' + nFormat(json.data[i].NewBu,2);
						if(txt.slice(-1) == '>') {
							txt += '0';
						}
					}
				}
				tr += '<td class="' + cls +'">' + txt + '</td>';
				//BU加工金額
				tr += '<td class="number">' + nFormat(json.data[i].Bu,2) + '</td>';
				//BU加工:請求
				cls = 'number';
				txt = nFormat(json.data[i].BillBuPrc,2);
				tr += '<td class="' + cls +'">' + txt + '</td>';
				//BU加工金額:請求
				tr += '<td class="number">' + nFormat(json.data[i].BillBu,2) + '</td>';
				//引取 秒
				txt = '';
				if (typeof json.data[i].HikitoriTm != "undefined") {
					txt = nFormat(json.data[i].HikitoriTm,0);
				}
				tr += '<td class="number">' + txt + '</td>';
				//引取 ＠
				txt = '';
				if (typeof json.data[i].HikitoriPrc != "undefined") {
					txt = nFormat(json.data[i].HikitoriPrc,2);
				}
				tr += '<td class="number">' + txt + '</td>';
				//引取 金額
				txt = '';
				if (typeof json.data[i].Hikitori != "undefined") {
					txt = nFormat(json.data[i].Hikitori,2);
				}
				tr += '<td class="number">' + txt + '</td>';
				tr += '</tr>';
			}
			$('#tbody').append(tr);
			tr = '<tr>'
			tr += '<td class="number">合計</td>';	// #
			tr += '<td></td>';		//受入日
			tr += '<td></td>';		//収単/担当者
			tr += '<td></td>';		//品番
			tr += '<td></td>';		//品名
//			tr += '<td></td>';		//受入数
			tr += '<td class="number">' + nFormat(json.data.reduce((a,x) => a += parseFloat(x.Qty), 0), 0) +'</td>';	//受入数
			tr += '<td></td>';		//工料
			tr += '<td class="number">' + nFormat(json.data.reduce((a,x) => a += parseFloat(x.Koryo),0),2) +'</td>';	//工料金額
			tr += '<td class="hakogai"></td>';		//箱代
			tr += '<td class="hakogai number">' + nFormat(json.data.reduce((a,x) => a += parseFloat(x.Hako),0),2) +'</td>';	//箱代金額
			tr += '<td class="hakogai"></td>';		//外装
			tr += '<td class="hakogai number">' + nFormat(json.data.reduce((a,x) => a += parseFloat(x.Gaiso),0),2) +'</td>';	//外装金額
			tr += '<td></td>';		//加工
			tr += '<td class="number">' + nFormat(json.data.reduce((a,x) => a += parseFloat(x.Kako),0),2) +'</td>';	//加工金額
			tr += '<td></td>';		//BU加工
			tr += '<td class="number">' + nFormat(json.data.reduce((a,x) => a += parseFloat(x.Bu),0),2) +'</td>';	//BU加工金額
			tr += '<td></td>';		//BU加工:請求
			tr += '<td class="number">' + nFormat(json.data.reduce((a,x) => a += parseFloat(x.BillBu),0),2) +'</td>';	//BU加工金額:請求
			txt = '';
			if (typeof json.data[0].HikitoriTm != "undefined") {
				txt = nFormat(json.data.reduce((a,x) => a += parseFloat(x.HikitoriTm),0),0);
			}
			tr += '<td class="number">' + txt +'</td>';	//引取 秒
			tr += '<td></td>';		//引取 ＠
			tr += '<td class="number">' + nFormat(json.data.reduce((a,x) => a += parseFloat(x.Hikitori),0),2) +'</td>';	//引取 金額
			tr += '</tr>';
			$('#table > tbody').append(tr);
//			$('#floatThead').trigger('change');
			$('#HakoGai').trigger('change');
			chime();
			$('#msg').text(i + '件 処理時間: ' + $('#lastTm').text() );
			$('#log').html($('#log').html() + json.rfbill + '<br>');
			$(".rfbill").attr("href", json.rfbill);
			$('.rfbill').show();
		});
        return false;
	});
	window.location.hash = '#tab_home';
	(function() {
		var	reinit = null;
	    var beforePrint = function() {
		    console.log('beforePrint().');
			$('#log').html($('#log').html() + '<br>beforePrint().');
			if( $('#floatThead').prop('checked') ) {
				reinit = $table.floatThead('destroy');
			}
	    };
	    var afterPrint = function() {
		    console.log('afterPrint().');
			$('#log').html($('#log').html() + '<br>afterPrint().');
//			$table.floatThead();
			if( reinit ) {
				reinit();
			}
//			fTable();
//			fTable().trigger('reflow');
//			$('#table').trigger('reflow');
	    };
	    if (window.matchMedia) {
		    var mediaQueryList = window.matchMedia('print');
		    mediaQueryList.addListener(function(mql) {
		        if (mql.matches) {
			        beforePrint();
		        } else {
		    	    afterPrint();
		        }
		    });
	    }
	    window.onbeforeprint = beforePrint;
	    window.onafterprint = afterPrint;
	}());
	// チェックボックスをチェックしたら発動
	$('#floatThead').change(function() {
	    console.log('floatThead.change():' + $(this).prop('checked') + ' ' + $(this).text());
		// もしチェックが入ったら
		if ($(this).prop('checked')) {
			$("#floatThead_text").text('固定する');
			$table.floatThead();
		} else {
			$("#floatThead_text").text('固定しない');
			$table.floatThead('destroy');
		}
	});
	// チェックボックスをチェックしたら発動
	$('#HakoGai').change(function() {
		// もしチェックが入ったら
		if ($(this).prop('checked')) {
			$("#HakoGai_text").text('表示');
			$('.hakogai').show();
		} else {
			$("#HakoGai_text").text('非表示');
			$('.hakogai').hide();
		}
	});
	if(!isIE()){
//		$('#floatThead').trigger('click');
		$('#floatThead').prop("checked",true);
		$('#floatThead').change();
	}
//	$('.drawer').drawer();
	$('.btn').on('click', function(){
		$('#floatThead').prop("checked",false);
		$('#floatThead').change();
	    var clipboard = new Clipboard('.btn');
	    clipboard.on('success', function(e) {
	        //成功時の処理
	    });
	    clipboard.on('error', function(e) {
	      //失敗時の処理
	    });
	});
	// 請求書(ダウンロード)
    $('.btn_bill').click(function () {
//text/plain
//application/vnd.ms-excel
//application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
		var blob = new Blob([ this.text ], { "type" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
		let link = document.createElement('a')
		link.href = window.URL.createObjectURL(blob)
		link.download = 'test.xlsx'
		link.click()
	});
});
function isIE() {
	var userAgent = window.navigator.userAgent.toLowerCase();
	if( userAgent.match(/(msie|MSIE)/) || userAgent.match(/(T|t)rident/) ) {
//	    var isIE = true;
		var ieVersion = userAgent.match(/((msie|MSIE)\s|rv:)([\d\.]+)/)[3];
		ieVersion = parseInt(ieVersion);
		$('#log').html('IE' + ieVersion);
		return true;
//	} else {
//	    var isIE = false;
	}
	return false;
}
/*
Chime音
*/
function chime(nm) {
	if(typeof nm === 'undefined') {
		nm = '#chime';
	}
	var $audio = $(nm).get(0);
	$audio.volume = 1.0;
	$audio.play();
}
function storage(n,v) {
	if(typeof v === 'undefined') {
		v = null;
	}
	var	stg = localStorage;
	if(v == null) {
		v = stg.getItem(n,v);
	} else {
		stg.setItem(n,v);
	}
	return v;
}
function dateTimeFormat(date,second) {
	if(typeof second === 'undefined') {
		second = null;
	}
  var y = date.getFullYear();
  var m = date.getMonth() + 1;
  var d = date.getDate();
  var w = date.getDay();
  var hh = date.getHours();
  var mm = date.getMinutes();
  var ss = date.getSeconds();
  var wNames = ['日', '月', '火', '水', '木', '金', '土'];

  m = ('0' + m).slice(-2);
  d = ('0' + d).slice(-2);
  hh = ('0' + hh).slice(-2);
  mm = ('0' + mm).slice(-2);
  ss = ':' + ('0' + ss).slice(-2);
  if (!second) {
	ss = '';
  }

  // フォーマット整形済みの文字列を戻り値にする
  return y + '.' + m + '.' + d + ' ' + hh + ':' + mm + ss;
}
$(document).ready(function() {
	var headerHight = 100; //ヘッダの高さ
	$('a[href^="#"]').click(function(){
		var href= $(this).attr("href");
		var target = $(href == "#" || href == "" ? 'html' : href);
		var position = target.offset().top-headerHight; //ヘッダの高さ分位置をずらす
		$("html, body").animate({scrollTop:0}, 550, "swing");
//		return false;
	});
});
