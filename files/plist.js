/*
plist.js
*/
var	arr = [];

arr.push('v0.08 2020.11.10 [Copy]ボタン：検索結果をクリップボードにコピー');
arr.push('v0.07 2020.04.30 編集でパレットNoが変更できないのを修正');
arr.push('v0.06 2020.04.06 No.がパレットNo単位でクリアされないのを修正');
arr.push('v0.05 2020.04.03 パレットNo範囲指定で荷積明細を印刷できるようにしました.');
arr.push('v0.04 2020.03.27 パレットNoクリックで印刷プレビューを表示するように変更');
arr.push('v0.03 2020.03.26 荷積明細 差異／区分を削除');
arr.push('v0.02 2020.03.26 明細書に品番番号バーコードを印字');
arr.push('v0.01 2020.03.04 滋賀dcエアコン用');

var	version = arr[0];
var	tHistory = arr.join('<p>');

function debug(t) {
	console.log('debug(): ' + t);
	$.toast({
		text : t	,
		loader: false	,
		hideAfter : 5 * 1000,
	});
}
function error(t) {
	console.log('error(): ' + t);
	$.toast({
		text : t		,
		hideAfter : 60 * 1000,
		icon : 'error'	,
	});
}
function log(t) {
	if(!t) {
		$('#log').html('');
	} else {
		$('#log').html($('#log').html() + t + '<br>');
	}
}
function getNow() {
	var now = new Date();
	var year = now.getFullYear();
	var mon = now.getMonth()+1; //１を足すこと
	var day = now.getDate();
	var hour = now.getHours();
	var min = now.getMinutes();
	var sec = now.getSeconds();
	//出力用
	return year + "/" + mon + "/" + day + "  " + padZero(hour) + ":" + padZero(min) + ":" + padZero(sec) + "";
}
//日時
function dateTimeFormat(date,second) {
	if(typeof second === 'undefined') {
		second = null;
	}
	if(isNaN(date.getFullYear())) {
		return '';
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
	log('');
	log('location.pathname: ' + location.pathname);
	log('location.pathname[0]: ' + location.pathname.split('/')[0]);
	log('location.pathname[1]: ' + location.pathname.split('/')[1]);
	log('location.basename: ' + location.pathname.split('/').pop());
	log('location.dirname: ' + location.pathname.replace(/\/[^/]*$/, ''));
	$('#key').text(location.pathname.replace(/\/[^/]*$/, '') + '/');

	$('a[href^="#"]').click(function() {
		$("html, body").animate({scrollTop:0}, 550, "swing");
	});
	//設定：初期値
	$("#dns").blur(function() {
//		if($(this).val() == '') {
//			$(this).val(location.pathname.split('/')[1]);
//		}
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
			case 'limit':	val = 300;	break;
			case 'dns':		val = location.pathname.split('/')[1];	break;
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
	// バージョン
	$('#msg').html(version);
	$('#version').html('');
	$('#version').html($('#version').html() + '<p>' + tHistory);
	// 検索ボタン
	$('#submit').on('click', function() {
		debug('■検索開始.' + this.id);
		var	req = 'plist.py?dns=' + $('#dns').val();
		req += '&pallet_no1=' + $('#pallet_no1').val();
		req += '&pallet_no2=' + $('#pallet_no2').val();
		req += '&limit=' + $('#limit').val();
		log(req);
		$(this).addClass("gif-load");
		fetch(req)
		.then((res) => {
			$(this).removeClass("gif-load");
			return res.json();
		})
		.then((json) => {
			// 
			var	tr = '';
			for ( var i = 0 ; i < json.list.length ; i++ ) {
				tr += '<tr>';
				tr += '<td class="number">' + (i + 1) + '</td>';
				tr += '<td class="ID_NO">' + json.list[i].KEY_ID_NO + '</td>';
				tr += '<td >' + json.list[i].KEY_HIN_NO + '</td>';
				tr += '<td class="qty">' + Number(json.list[i].SURYO) + '</td>';
				tr += '<td class="pallet_no">';
				tr += '<a href="#pallect_list" class="plist">' + json.list[i].LK_SEQ_NO + '</a>';
				tr += '</td>';
				tr += '<td class="case_qty">' + Number(json.list[i].KENPIN_SURYO) + '</td>';
				tr += '<td><input type="button" value="編集" class="btn_edit"></td>';
				tr += '<td >' + json.list[i].KENPIN_YMD + ' ' + json.list[i].KENPIN_HMS.substr( 0, 4 ) + '</td>';
				tr += '<td class="qty">' + json.list[i].Row + '</td>';
				tr += '<td class="text">' + json.list[i].Biko1 + '</td>';
				tr += '<td class="text center">' + json.list[i].KanDt + '</td>';
				tr += '<td class="text center">' + json.list[i].KanQty + '</td>';
				tr += '<td class="text center">' + json.list[i].NaraDt + '</td>';
				tr += '</tr>';
			}
			tr += '<tr><td colspan="' + $('#list thead tr').find("th").length + '">';
			tr += json.list.length + '件';
			if (json.list.length >= Number($('#limit').val())) {
				tr += ' ※最大件数に達しました.';
			}
			tr += '</td>';
			$('#list tbody').find("tr").remove();
			$('#list tbody').append(tr);
		})
		.catch(function(err) {
			console.log('catch:' + err);
			var tr = '';
			tr += '<tr><td colspan="' + $('#list thead tr').find("th").length + '">';
			tr += req + '<br>';
			tr += err;
			tr += '</td>';
			$('#list tbody').find("tr").remove();
			$('#list tbody').append(tr);
		});
		return false;
	});
	// 出庫表－登録
	$(document).on('click', '.btn_edit', function() {
		console.log('.btn_edit:click' + this.id);
//		$("#e_pallet_no").val($(this).closest('tr').children("td.pallet_no").text());
//		$("#ID_NO").val($(this).closest('tr').children("td.ID_NO").text());
		var	tr = $(this).closest('tr');
		$("#e_pallet_no").val(tr.children("td.pallet_no").text());
		$("#e_case_qty").val(tr.children("td.case_qty").text());
		$("#ID_NO").text(tr.children("td.ID_NO").text());
		$("#Row").text(tr[0].rowIndex);
		console.log('#e_pallet_no:' + $("#e_pallet_no").val());
		console.log('#ID_NO:' + $("#ID_NO").text());
		$("#pallet_edit").dialog('option', 'title', 'パレットNo：編集 - ' + $("#Row").text());
        $("#pallet_edit").dialog("open");
		return false;
	});
    $("#pallet_edit").dialog({
        autoOpen: false,
        height: 'auto',
        width: 'auto',
        modal: true,
		title: 'パレットNo：編集',
        buttons: {  // ダイアログに表示するボタンと処理
			"変更": function() {
				//パレットNo変更
				$('#list tbody tr').eq($("#Row").text()-1).children('td.pallet_no').children('a').text($("#e_pallet_no").val());
				$('#list tbody tr').eq($("#Row").text()-1).children('td.case_qty').text($("#e_case_qty").val());

				var	url = 'plist.py?dns=' + $('#dns').val();
				url += '&id_no=' + $('#ID_NO').text();
				url += '&case_qty=' + $('#e_case_qty').val();
				url += '&pallet_no1=' + $('#e_pallet_no').val();
				console.log(url);
//				$("#pallet_edit .msg").text("処理中...");
//				$("#pallet_edit .msg").addClass("gif-load");
				fetch(url)
				.then((res) => {
					console.log(res);
					return res.json();
				})
				.then((json) => {
					console.log(json);
					$.toast({
						text : 'パレットNo変更OK.'	,
						loader: false			,
						hideAfter : 5 * 1000	,
					});
//					$(this).dialog("close");
				})
				.catch(function(err) {
					console.log('catch:' + err);
					alert(err);
				});
				$(this).dialog("close");
			},
            "キャンセル": function() {
				$(this).dialog("close");
            }
        },
        // ダイアログのイベント処理
        open: function(event, ui) {
			console.log('dialog.open():' + this.id);
        },
        close: function() {
			console.log('dialog.close():' + this.id);
        }
	});
	// 荷積明細印刷
	$(document).on('click', 'a.plist', function() {
		console.log('click a.plist:' + this.id + ' class=' + $(this).attr("class"));
		var	tr = $(this).closest('tr');
		$('#pallet_no1').val($(this).closest('tr').children("td.pallet_no").text());
		$('#pallet_no2').val('');
		$('#plist').trigger('click');
		return false;
	});
	// 荷積明細印刷ボタン（複数）
	$(document).on('click', '#plist', function() {
		console.log('click :' + this.id + ' class=' + $(this).attr("class"));
		var	req = 'plist.py?dns=' + $('#dns').val();
		req += '&pallet_no1=' + $('#pallet_no1').val();
		req += '&pallet_no2=' + $('#pallet_no2').val();
		req += '&limit=' + $('#limit').val();
		$(this).addClass("gif-load");
		var	title = '荷積明細-' + $('#pallet_no1').val() + ' - ' +$('#pallet_no2').val();
		var	html = '<html><head><title>' + title + '</title>';
		html += '<link type="text/css" rel="stylesheet" href="plist_p.css?' + getCurrentTime() + '">';
		html += '</head><body>';
		html += title + '<span class="gif-load">...検索中</span>';
		html += '</body></html>';
		var	w = window.open("", title);
		w.document.write(html);
		w.document.close();
		fetch(req)
		.then((res) => {
			$(this).removeClass("gif-load");
			return res.json();
		})
		.then((json) => {
			// 荷積明細－table
			var	tr = '';
			var	iCnt = 0;
			var	tQty = 0;
			var	tCase = 0;
			for ( var i = 0 ; i < json.list.length ; i++ ) {
				if (i == 0 || json.list[i].LK_SEQ_NO != json.list[i-1].LK_SEQ_NO) {
					tr += '<div class="center noprint">';
					tr += '<a href="#" onclick="window.print(); return false;">印刷する</a>';
					tr += '　';
					tr += '<a href="#" onclick="window.close(); return false;">閉じる</a>';
					tr += '</div>';
					tr += '<div class="hd_now">' + getNow() + '</div>';
					tr += '<div class="hd_title">積込・荷卸 明細書</div>';
					tr += '<table class="head">';
					tr += '<tr>';
					tr += '<td class="prof">行き先</td><td class="prof">甲西→奈良</td>';
					tr += '<td>　　　　　　　　</td>'
					tr += '<td rowspan="2" class="barcode">'
					tr += '<div class="">*' + json.list[i].LK_SEQ_NO + '*</div><div class="text">' + json.list[i].LK_SEQ_NO + '</div>';
					tr += '</td>';
					tr += '</tr>';
					tr += '<tr>';
					tr += '<td class="prof">パレットＮｏ.</td><td class="prof">' + json.list[i].LK_SEQ_NO + '</td>';
					tr += '</tr>';
					tr += '</table>';
					tr += '<table class="list"><tbody>';
					tr += '<thead>';
					tr += '<tr>';
					tr += '<th>No.</th>';
					tr += '<th>品目番号</th>';
					tr += '<th>品目番号バーコード</th>';
		//			tr += '<th>収支</th>';
		//			tr += '<th>相手先</th>';
					tr += '<th>数量</th>';
					tr += '<th>ID-No.</th>';
					tr += '<th>種別</th>';
					tr += '<th>箱数</th>';
					tr += '<th>荷卸数量</th>';
		//			tr += '<th>差異</th>';
		//			tr += '<th>区分</th>';
					tr += '</tr>';
					tr += '</thead>';
				}
				tQty += Number(json.list[i].SURYO);
				tCase += Number(json.list[i].KENPIN_SURYO);
				iCnt++;
				tr += '<tr>';
				tr += '<td class="number">' + iCnt + '</td>';
				tr += '<td class="pn">' + json.list[i].KEY_HIN_NO + '</td>';
				tr += '<td class="pn-barcode">*' + json.list[i].KEY_HIN_NO + '*</td>';
//				tr += '<td >' + json.list[i].SYUKO_SYUSI + '</td>';
//				tr += '<td >' + json.list[i].KEY_MUKE_CODE + '</td>';
				tr += '<td class="qty">' + Number(json.list[i].SURYO) + '</td>';
				tr += '<td class="">' + json.list[i].KEY_ID_NO + '</td>';
				tr += '<td class="">受注</td>';
				tr += '<td class="case_qty">' + Number(json.list[i].KENPIN_SURYO) + '</td>';
				tr += '<td class=""></td>';
//				tr += '<td class=""></td>';
//				tr += '<td class="">1</td>';
				tr += '</tr>';
				if((i + 1 ) >= json.list.length || json.list[i].LK_SEQ_NO != json.list[i+1].LK_SEQ_NO) {
					tr += '<tr class="total">';
					tr += '<td ></td>';
					tr += '<td ></td>';
					tr += '<td >合計</td>';
		//			tr += '<td ></td>';
					tr += '<td class="qty">' + tQty + '</td>';
					tr += '<td ></td>';
					tr += '<td ></td>';
					tr += '<td class="case_qty">' + tCase + '</td>';
					tr += '<td ></td>';
		//			tr += '<td ></td>';
		//			tr += '<td ></td>';
					tr += '</tr>';
					tr += '</tbody></table>';
					tQty = 0;
					tCase = 0;
					iCnt = 0;
					if((i + 1 ) < json.list.length) {
						tr += '<div class="pagebreak"></div>'
					}
				}
			}
			var	title = '荷積明細-' + json.list[0].LK_SEQ_NO;
			var	html = '<html><head><title>' + title + '</title>';
			html += '<link type="text/css" rel="stylesheet" href="plist_p.css?' + getCurrentTime() + '">';
			html += '</head><body onLoad="window.print();">';
			html += tr;
			html += '</body></html>';
			w.document.clear();
			w.document.write(html);
			w.document.close();
//			w.print();
//			w.close();
		});
		return false;
	});
	$('.btn').on('click', function(){
	    var clipboard = new Clipboard('.btn');
	    clipboard.on('success', function(e) {
	        //成功時の処理
	    });
	    clipboard.on('error', function(e) {
	      //失敗時の処理
	    });
	});
	window.location.hash = '#tab_home';
//	$('#submit').trigger('click');
//	debug("ページ読込完了");
});
//現在時刻取得（yyyymmddhhmmss）
function getCurrentTime() {
    var now = new Date();
    var res = "" + now.getFullYear() + padZero(now.getMonth() + 1) + padZero(now.getDate()) + padZero(now.getHours()) + 
        padZero(now.getMinutes()) + padZero(now.getSeconds());
    return res;
}

//先頭ゼロ付加
function padZero(num) {
    return (num < 10 ? "0" : "") + num;
}
function plist_head(y_syuka, y_syuka0) {
	var	tr = '';
	tr += '<table>';
	tr += '<thead>';
	tr += '<tr>';
	tr += '<td colspan="4">';
	tr += '積水 出庫表 ' + y_syuka.KEY_SYUKA_YMD;
	tr += ' ' + y_syuka.DEN_NO;
	if(y_syuka0 == 'y_syuka0') {
		tr += ' (未出庫)';
	}
	tr += '</td>';
	tr += '<td colspan="2" class="right">';
	tr += y_syuka.KEY_MUKE_CODE;
	tr += '　' + y_syuka.MUKE_NAME;
	tr += '<span>';
	tr += '　　' + y_syuka.BIKOU1;
	tr += '</span>';
	tr += '</td>';
	tr += '</tr>';
	tr += '<tr>';
	tr += '<th>#</th>';
	tr += '<th>標準棚番</th>';
	tr += '<th>品番</th>';
	tr += '<th>受注数</th>';
	tr += '<th>ID-No</th>';
	tr += '<th></th>';
	tr += '</tr>';
	tr += '</thead>';
	tr += '<tbody>';
	return tr;
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
