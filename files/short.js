/*
short.js
*/
// バージョン
var	ver = [];
ver.push('v0.09 2021.04.15 収支(SJ)追加');
ver.push('v0.08 2021.04.15 Excel出力 入荷予定２件→５件に変更');
ver.push('v0.07 2021.04.15 入荷予定を２件→３件に変更');
ver.push('v0.06 2021.04.14 滋賀pc エアコン(A)対応');
ver.push('v0.05 2020.07.01 4:炊飯／5:調理小物の入荷予定(商品化計画 入荷予定)も参照するように変更');
ver.push('v0.04 2020.06.17 小野pc対応');
ver.push('v0.03 2020.04.13 操作性向上');
ver.push('v0.02 2020.03.24 処理中gif変更');
ver.push('v0.01 2020.03.19 滋賀pc用');
$('#msg').html(ver[0]);
$('#version').html(ver.join('<p>'));

function log(t) {
	if(!t) {
		$('#log').html('');
	} else {
		$('#log').html($('#log').html() + t + '<br>');
	}
}
$(document).ready(function() {
	$('a[href^="#"]').click(function() {
		$("html, body").animate({scrollTop:0}, 550, "swing");
//		return false;
	});
	//Table更新日時セット
	$('span.table').on('click', function() {
		log('parent:' + $(this).parent().parent().parent().attr('id'));
		log($(this).text());
		log('th:' + $($(this).parent(),'thead tr').find("th").length);
		log('th:' + $(this).parent().find('thead tr').find("th").length);

		var	req = 'dbinfo.py?dns=' + $('#dns').val();
		if($('#dbinfo').val()) {
			req = $('#dbinfo').val();
		}
//		req += '&table=' + $(this).text().split(':')[0];
		req += '&table=' + $(this).attr('title');
//		if($('#files').val()) {
//			req += '&files=' + $('#files').val();
//		}
		log(req);
		$(this).addClass("gif-load");
		fetch(req)
		.then((res) => {
			$(this).removeClass("gif-load");
			log(res.headers.get("content-type"));
			return res.json();
		})
		.then((json) => {
			log(JSON.stringify(json));
//			$(this).text(json.table[0].Xf$Name + ':更新日時 ' + json.table[0].mtime);
			$(this).text('更新日時 ' + json.table[0].mtime);
		});
	});
	//前回処理日時セット
	$('.result').on('click', function() {
		var	pid = $(this).closest('.tab_content').attr('id');
		$(this).text('...' + pid);
		var	req = '';
		switch(pid) {
		case 'tab_gorder':	req = 'gorder.py';		break;
		case 'tab_pnshort':	req = 'pnshort.py';		break;
		case 'tab_hosoku':	req = 'shortxls.py';	break;
		default:	return;							break;
		}
		req += '?file=log';
		log(req);
		$(this).addClass("gif-load");
		fetch(req)
		.then((res) => {
			$(this).removeClass("gif-load");
			log(res.headers.get("content-type"));
			$(this).closest('.upload').attr('id');
			return res.json();
		})
		.then((json) => {
			log(JSON.stringify(json));
			$(this).text('(前回) ' + json.log + '');
			$(this).closest('.upload').files[0].name = json.log;
		});
	});
	$('.result').on('change', function() {
		$(this).hide().fadeIn('slow');
	});
	$('.result').trigger('click');
	//テーブル 検索
	$('.search').on('click', function() {
		log('click():' + $(this).val());
		log($(this).parent().parent().parent().attr('id'));
		var	req = '';
		switch($(this).parent().parent().parent().attr('id')) {
		case 'tab_gorder':	req = 'gorder.py';		break;
		case 'tab_pnshort':	req = 'pnshort.py';		break;
		case 'tab_hosoku':	req = 'shortxls.py';	break;
		}
		req += '?dns=' + $('#dns').val();
		req += '&limit=' + $('#limit').val();
		log(req);
		$(this).addClass("gif-load");
		fetch(req)
		.then((res) => {
			$(this).removeClass("gif-load");
			return res.json();
		})
		.then((json) => {
			$(this).closest('.content').addClass("gif-load");
			log(json.table);
			log(json.log);
			log(json.list.length);
			var	tr = '';
			for ( var i = 0 ; i < json.list.length ; i++ ) {
				tr += '<tr>';
				switch(json.table) {
				case 'GOrder':
					tr += '<td class="number">' + json.list[i].Row + '</td>';
					tr += '<td class="center">' + json.list[i].IdNo + '</td>';
					tr += '<td class="pn">' + json.list[i].Pn + '</td>';
					tr += '<td class="qty">' + json.list[i].Qty + '</td>';
					tr += '<td class="center">' + json.list[i].YoteiDt + '</td>';
					tr += '<td class="center">' + json.list[i].OrderDt + '</td>';
					tr += '<td class="center">' + json.list[i].ShiteiDt + '</td>';
					tr += '<td class="center">' + json.list[i].KaitoDt + '</td>';
					break;
				case 'PnShort':
					tr += '<td class="number">' + (i + 1) + '</td>';
					tr += '<td class="text">' + json.list[i].JCode + '</td>';
					tr += '<td class="pn">' + json.list[i].Pn + '</td>';
					tr += '<td class="text">' + json.list[i].PName + '</td>';
					tr += '<td class="text">' + json.list[i].DModel + '</td>';
					tr += '<td class="text">' + json.list[i].DModelName + '</td>';
					tr += '<td class="text">' + json.list[i].PnCate + '</td>';
					tr += '<td class="text">' + json.list[i].PnCateName + '</td>';
					tr += '<td class="text">' + json.list[i].JpSt + '</td>';
					tr += '<td class="text">' + json.list[i].JpEn + '</td>';
					tr += '<td class="text">' + json.list[i].JpKb + '</td>';
					tr += '<td class="text">' + json.list[i].OsSt + '</td>';
					tr += '<td class="text">' + json.list[i].OsEn + '</td>';
					tr += '<td class="text">' + json.list[i].OsKb + '</td>';
					tr += '<td class="text">' + json.list[i].Biko + '</td>';
					break;
				case 'ShortXls':
					tr += '<td class="number">' + (i + 1) + '</td>';
					tr += '<td class="pn">' + json.list[i].Pn + '</td>';
					tr += '<td class="text">' + json.list[i].Biko + '</td>';
					tr += '<td class="text">' + json.list[i].PnBiko + '</td>';
					break;
				}
				tr += '</tr>';
			}
//			tr += '<tr><td colspan="' + $('#tab_gorder thead tr').find("th").length + '">';
			tr += '<tr><td></td><td colspan="' + ($($(this).closest('.content'),'thead tr').find("th").length - 1) + '">';
			tr += json.list.length + '件';
			if (json.limit > 0 && json.list.length >= json.limit) {
				tr += ' ※最大件数に達しました.';
			}
			tr += '</td>';
//			$('#tab_gorder tbody').find("tr").remove();
//			$('#tab_gorder tbody').append(tr);
			$(this).closest('.content').find('tbody').find("tr").remove();
			$(this).closest('.content').find('tbody').append(tr);
			$(this).closest('.content').removeClass("gif-load");
			chime();
		}).catch((error) => {
			log(error);
			$(this).next('.result').text(error);
			chime('#chime3');
		});
	});
	//アップロード
	$('.upload').on('change', function() {
		log($(this).attr('class') + ':' + this.files[0].name);
		var	pid = $(this).closest('.tab_content').attr('id');
		var	req = '';
		switch(pid) {
		case 'tab_gorder':	req = 'gorder.py';		break;
		case 'tab_pnshort':	req = 'pnshort.py';		break;
		case 'tab_hosoku':	req = 'shortxls.py';	break;
		}
		req += '?dns=' + $('#dns').val();
		log(req);
		var formData = new FormData();
		formData.append('upload', this.files[0]);
		$(this).parent().addClass("gif-load");
		$(this).next('.result').text(this.files[0].name + '...処理中' );
		fetch(req, {
			method: 'POST',
			body: formData
		}).then((res) => {
			$(this).val('');
			$(this).parent().removeClass("gif-load");
			return res.json();
		}).then((json) => {
			log(json);
//			$('#tab_gorder .result').text('完了しました.');
//			$(this).next('.uploadValue').val('完了しました.');
//			$('#tab_gorder .table').trigger('click');
//			var	txt = this.files[0].name + ' ' + json.load.length + '件 ' + json.mtime + '.OK';
			var	txt = json.log + '.OK';
			$(this).parent().next('.result').fadeOut('slow',function(){
				$(this).text(txt);
				$(this).fadeIn('slow',function(){
				});
			});
			$(this).parent().parent().find('.table').trigger('click');
			chime();
//		}).catch(function(error) {
		}).catch((error) => {
			log(error);
//			$('#tab_gorder .result').text(error);
			$(this).next('.result').text(error);
//			$(this).next('.uploadValue').val(error);
			chime('#chime3');
		});
	});
	//Excel
	$('input.excel').on('click', function() {
//		alert("input.excel");
		log('■Excel.' + this.id);
		var	req = 'short.py?dns=' + $('#dns').val();
		req += '&excel=1';
		req += '&jgyobu=' + $('#jgyobu').val().toUpperCase();
		log(req);
		$(this).addClass("gif-load");
		fetch(req)
		.then((res) => {
			$(this).removeClass("gif-load");
			return res.json();
		})
		.then((json) => {
			log(json);
			location.href = json.excel;
			chime();
		}).catch((error) => {
			log(error);
			chime('#chime3');
		});
	});
	// 検索ボタン
	$('#submit').on('click', function() {
		log('■検索開始.' + this.id);
		var	req = 'short.py?dns=' + $('#dns').val();
		req += '&limit=' + $('#limit').val();
		req += '&jgyobu=' + $('#jgyobu').val().toUpperCase();
		log(req);
		$(this).addClass("gif-load");
		fetch(req)
		.then((res) => {
			$(this).removeClass("gif-load");
			return res.json();
		})
		.then((json) => {
			var	tr = '';
			var now = new Date();
			var	year = now.getFullYear() + '-';
			for ( var i = 0 ; i < json.list.length ; i++ ) {
				tr += '<tr>';
				tr += '<td class="number">' + (i + 1) + '</td>';
				tr += '<td class="pn">' + json.list[i].HIN_GAI + '</td>';
				tr += '<td class="pn">' + json.list[i].HIN_NAI + '</td>';
				tr += '<td class="text">' + json.list[i].HIN_NAME + '</td>';
				tr += '<td class="center">' + json.list[i].Nai + '</td>';
				tr += '<td class="center">' + json.list[i].Gai + '</td>';
				tr += '<td class="qty">' + json.list[i].qty + '</td>';
				tr += '<td class="qty">' + json.list[i].AveSyuka + '</td>';
				tr += '<td class="qty">' + json.list[i].ZMonth + '</td>';
				tr += '<td class="center">' + json.list[i].Month1 + '</td>';
				tr += '<td class="center">' + json.list[i].Day5 + '</td>';
				tr += '<td class="center">' + json.list[i].NaiDisconYm + '</td>';
				tr += '<td class="text">' + json.list[i].Biko + '</td>';
				tr += '<td class="text">' + json.list[i].PnxBiko + '</td>';
				tr += '<td class="qty">' + ((json.list[i].jp == 0) ? '' : json.list[i].jp) + '</td>';
				tr += '<td class="qty">' + ((json.list[i].s8 == 0) ? '' : json.list[i].s8) + '</td>';
				tr += '<td class="qty">' + ((json.list[i].sj == 0) ? '' : json.list[i].sj) + '</td>';
				tr += '<td class="qty">' + ((json.list[i].other == 0) ? '' : json.list[i].other) + '</td>';
				tr += '<td class="center" title="' + json.list[i].YoteiDt + ' ' +  json.list[i].Qty + '">'
				tr += json.list[i].YoteiDt.replace(year,'') + '<br>' + json.list[i].Qty + '</td>';
				tr += '<td class="center" title="' + json.list[i].YoteiDt2 + ' ' +  json.list[i].Qty2 + '">'
				tr += ((json.list[i].YoteiDt2 == null) ? '' : json.list[i].YoteiDt2.replace(year,'') + '<br>' + json.list[i].Qty2) + '</td>';
				tr += '<td class="center" title="' + json.list[i].YoteiDt3 + ' ' +  json.list[i].Qty3 + '">'
				tr += ((json.list[i].YoteiDt3 == null) ? '' : json.list[i].YoteiDt3.replace(year,'') + '<br>' + json.list[i].Qty3) + '</td>';
//				tr += '<td class="center">' + ((json.list[i].Qty2 == null) ? '' : json.list[i].Qty2) + '</td>';
				tr += '</tr>';
			}
			tr += '<tr><td colspan="' + $('#list thead tr').find("th").length + '">';
			tr += json.list.length + '件';
			if (Number($('#limit').val()) > 0 && json.list.length >= Number($('#limit').val())) {
				tr += ' ※最大件数に達しました.';
			}
			tr += '</td>';
			$('#list tbody').find("tr").remove();
			$('#list tbody').append(tr);
			chime();
		}).catch((err) => {
			log('catch:' + err);
			var tr = '';
			tr += '<tr><td colspan="' + $('#list thead tr').find("th").length + '">';
			tr += req + '<br>';
			tr += err;
			tr += '</td>';
			$('#list tbody').find("tr").remove();
			$('#list tbody').append(tr);
			chime('#chime3');
		});
		return false;
	});
	window.location.hash = '#tab_home';
//	$('#submit').trigger('click');
	$('.table').trigger('click');
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
