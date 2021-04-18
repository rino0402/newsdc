// バージョン
var	ver = [];
var	ver007 = '0.07 2021.04.12 項目並び変更：所定内 残業 深夜 休出 有休日 有休時間 遅刻 早退 備考<br>';
ver007 += '・休出(休日出勤)を追加<br>';
ver007 += '・出勤日数＝始業、有休<br>';
ver007 += '・所定内時間→実労働時間<br>';
ver007 += '・合計時間＝有休＋実労働＋残業＋深夜＋休出<br>';
ver007 += '・勤務表の項目変更対応<br>';
ver007 += '2021.04.19 対応予定<br>';
ver007 += '・各自の合計行を追加<br>';
ver007 += '2021.04.26 対応予定<br>';
ver007 += '・チェック済欄追加<br>';
ver.push(ver007);
ver.push('0.06 2021.03.25 修正：始業なしの場合、出勤日数をカウントしないように修正');
ver.push('0.05 2021.03.24 変更：従業員で退職者を除く(選択可)ように変更');
ver.push('0.04 2021.03.23 変更：出勤にしない場合はシフト"--"を選択して下さい。');
ver.push('0.03 2021.03.22 修正：所定内がマイナスになる不具合を修正。');
ver.push('0.02 2021.03.18 変更：有休時間 を追加しました。');
ver.push('0.01 2021.03.16 修正：出勤/退勤の両方がない場合、始業／終業が入力できない。');
$('#version').html(ver.join('<p>'));
$('#msg').html(ver[0]);

function fmtDay(d) {
	return d.getMonth() + '/' + d.getDay();
}
function fmtHour(d) {
//	console.log('fmtHour(' + d + ')');
//	console.log('(d | 0)=' + (d | 0));
	if ( Number.isFinite(d) ) {
		if ( d != 0 ) {
			return d.toFixed(2);
		}
	}
	return '';
}
function fmtTm(tm) {
	if(tm) {
//		console.log(tm + '=' + tm.prototype.toString);
//		console.log(tm + '=' + typeof tm);	//08:45:00=string
//		return tm.getHours() + ':' + tm.getMinutes();
		return tm.slice(0,5);
	} else {
		return '';
	}
}
function fmtDt(dt) {
	if(dt) {
//		console.log(tm + '=' + tm.prototype.toString);
//		console.log(dt + '=' + typeof dt);	//08:45:00=string
//		return tm.getHours() + ':' + tm.getMinutes();
		return dt.slice(0,10);
	} else {
		return '';
	}
}

//テーブル 検索
$("input[name='load']").on('click', function() {
	var	req = 'load.py?dns=' + $('#dns').val();
	req += '&month=' + $("input[name='month']").val();
	if($("input[name='post']").val()) {
		req += '&post=' + $("input[name='post']").val();
	}
//	$(this).addClass("gif-load");
	dispLoading('検索中');
	$('#msg').text(req);
	fetch(req).then((res) => {
		$('#msg').text('');
//		$(this).removeClass("gif-load");
		removeLoading();
		return res.json();
	}).then((json) => {
		var	tr = '';
//		tr = '<tr>';
//		for (var item in json.columns) {
//				tr += '<td>' + json.columns[item][1] + '</td>';
//			tr += '<td>' + json.columns[item] + '</td>';
//		}
//		tr += '</tr>';
		for ( var i = 0 ; i < json.data.length ; i++) {
			var	id = json.data[i]['StaffNo'] + '_' + json.data[i]['strDt'];
			var	cls = '';
//			if (json.data[i].strDt.slice(-2) == '16') {
			if (i == 0 || json.data[i]['StaffNo'] != json.data[i-1]['StaffNo']) {
				cls = 'dt-top';
			}
			tr += '<tr id="' + id + '" class="' + cls + '">';
			tr += '<td>' + id + '</td>';
			var	post = json.data[i].Post;
			if (post) {
				post = post.trim();
			}
			tr += '<td class="Post">' + post + '</td>';
			tr += '<td class="StaffNo">' + json.data[i]['StaffNo'] + '</td>';
			tr += '<td class="Name">' + json.data[i]['Name'] + '</td>';
			var	cls = '';
			if (json.data[i]['Holiday'] != '') {
				cls = ' holiday';
			}
			tr += '<td class="date ' + json.data[i]['strDay'] + cls + '" title="' + json.data[i]['Holiday'] + '">' + json.data[i]['fmtDt'] + '</td>';
			var	shift = json.data[i].Shift;
			if( shift == '00') {
				shift = '休出';
			}
			tr += '<td class="Shift">' + shift + '</td>';
			if (json.data[i]['BegTm_i'] == '') {
				tr += '<td class="time BegTm">' + json.data[i]['BegTm5'] + '</td>';
			} else {
				tr += '<td class="time BegTm modify" title="' + json.data[i]['BegTm5'] + '">';
				tr += json.data[i]['BegTm_i'] + '</td>';
			}
			if (!json.data[i]['StartTm_i']) {
				tr += '<td class="time StartTm">' + fmtTm(json.data[i]['StartTm']) + '</td>';
			} else {
				tr += '<td class="time StartTm modify" title="' + fmtTm(json.data[i]['StartTm']) + '">';
				tr += fmtTm(json.data[i]['StartTm_i']) + '</td>';
			}
			if (!json.data[i]['FinishTm_i']) {
				tr += '<td class="time FinishTm">' + fmtTm(json.data[i]['FinishTm']) + '</td>';
			} else {
				tr += '<td class="time FinishTm modify" title="' + fmtTm(json.data[i]['FinishTm']) + '">';
				tr += fmtTm(json.data[i]['FinishTm_i']) + '</td>';
			}
			if (json.data[i]['FinTm_i'] == '') {
				tr += '<td class="time FinTm">' + json.data[i]['FinTm5'] + '</td>';
			} else {
				tr += '<td class="time FinTm modify" title="' + json.data[i]['FinTm5'] + '">';
				tr += json.data[i]['FinTm_i'] + '</td>';
			}
			if (json.data[i]['Actual_i'] || json.data[i]['Actual_i'] == 0) {
				tr += '<td class="hour Actual modify" title="' + (fmtHour(json.data[i]['Actual']) || '0.00') + '">';
				tr += fmtHour(json.data[i]['Actual_i']) || '0.00' + '</td>';
			} else {
				tr += '<td class="hour Actual">';
				tr += fmtHour(json.data[i]['Actual']) + '</td>';
			}
			if (json.data[i]['Extra_i'] || json.data[i]['Extra_i'] == 0) {
				tr += '<td class="hour Extra modify" title="' + (fmtHour(json.data[i]['Extra']) || '0.00') + '">';
				tr += fmtHour(json.data[i]['Extra_i'])  || '0.00' + '</td>';
			} else {
				tr += '<td class="hour Extra">' + fmtHour(json.data[i]['Extra']) + '</td>';
			}
			if (json.data[i]['Night_i'] || json.data[i]['Night_i'] == 0) {
				tr += '<td class="hour Night modify" title="' + (fmtHour(json.data[i]['Night']) || '0.00') + '">';
				tr += fmtHour(json.data[i]['Night_i']) || '0.00' + '</td>';
			} else {
				tr += '<td class="hour Night">' + fmtHour(json.data[i]['Night']) + '</td>';
			}
			if (json.data[i]['Dayoff_i'] || json.data[i]['Dayoff_i'] == 0) {
				tr += '<td class="hour Dayoff modify" title="' + (fmtHour(json.data[i]['Dayoff']) || '0.00') + '">';
				tr += fmtHour(json.data[i]['Dayoff_i']) || '0.00' + '</td>';
			} else {
				tr += '<td class="hour Dayoff">' + fmtHour(json.data[i]['Dayoff']) + '</td>';
			}
			tr += '<td class="hour PTO">' + fmtHour(json.data[i]['PTO']) + '</td>';
			tr += '<td class="hour PTO_tm">' + fmtHour(json.data[i]['PTO_tm']) + '</td>';
			tr += '<td class="hour">' + fmtHour(json.data[i]['Late']) + '</td>';
			tr += '<td class="hour">' + fmtHour(json.data[i]['Early']) + '</td>';
			tr += '<td class="Memo">' + json.data[i]['Memo'] + '</td>';
			tr += '<td class="total Days"></td>';
			tr += '<td class="total PTO_H"></td>';
			tr += '<td class="total Actual_H"></td>';
			tr += '<td class="total Extra_H"></td>';
			tr += '<td class="total Night_H"></td>';
			tr += '<td class="total Dayoff_H"></td>';
			tr += '<td class="total Total_H"></td>';
//			for (var item in json.data[i]) {
//				tr += '<td>' + json.data[i][item] || '' + '</td>';
//			}
			tr += '</tr>';
		}
		$('#list tbody').find("tr").remove();
		$('#list tbody').append(tr);
		$('#list').exTableFilter('#filter',
								{ignore : '0,1,5,6,7,8,9,10,11,12,13,14,15,16,17,19,20,21,22,23,24,25'
								,elementAutoBindTrigger : 'change'
								}
								);
		$('#filter').trigger('change');	//フィルター実行
		$('#tab_month input[name="edit"]').trigger('click');	//編集可
		total();
	}).catch((error) => {
		$('#msg').text(error);
	});
});
var	shift_list =  '';	//'{"00": "休出", "04": "04", "09": "09"}';
//[{"01":"いろはにほへと"},{"02":"ちりぬるを"},{"11":"わかよたれそ"},{"12":"つねならむ"}]
$(document).ready(function() {
	var	req = 'shift.py?dns=' + $('#dns').val();
	console.log(req);
	fetch(req).then((res) => {
		return res.json();
	}).then((json) => {
		console.log(json);
		shift_list = '{"--_": "--", "00_": "休出"';
		for ( var i = 0 ; i < json.length ; i++) {
			shift_list += ',"' + json[i].Shift + '_": "' + json[i].Shift + '"';
		}
		shift_list += '}';
	}).catch(function(err) {
		console.log(req + ' error:' + err);
    });
});
$('#tab_month input[name="edit"]').on('click', function(){
		console.log('edit:click');
		console.log('shift_list=' + shift_list);
		console.log(jQuery.parseJSON(shift_list));
		$('#log').html($('#log').html() + '<br>' + 'shift_list=' + shift_list);
		$('#list').Tabledit({
			url: 'edit.py?dns=' + $('#dns').val(),
			editButton: false,
			deleteButton: false,
			hideIdentifier: true,
			onDraw: function () {
				console.log('onDraw()');
			},
			onAjax: function(action, serialize) {
				//Ajax開始時
				console.log('onAjax(action, serialize)');
				console.log(action);
				console.log(serialize);
				var urlParams = new URLSearchParams(serialize);
				console.log(urlParams.get('id'));
				var id = '#' + urlParams.get('id');
				if(urlParams.get('PTO')) {
					console.log('PTO text=' + $(id).children("td.PTO").text());
					console.log('PTO html=' + $(id).children("td.PTO").html());
					console.log('PTO <span> text=' + $(id).find("td.PTO span").text());
					console.log('PTO <span> html=' + $(id).find("td.PTO span").html());
					console.log('PTO <input> val()=' + $(id).find("td.PTO input").val());
//					$(id).find("td.PTO span").text('10');
//					$(id).find("td.PTO input").val('10');
//					$(id).children("td.PTO").text(fmtHour(parseFloat(urlParams.get('PTO'))));
				}
//				return false;
				$(id).addClass("gif-load");
			},
			onSuccess: function(data, textStatus, jqXHR) {
				console.log('onSuccess(data, textStatus, jqXHR)');
				console.log(data);
				console.log(textStatus);
				console.log(jqXHR);
				console.log(data.id);
				var id = '#' + data.id;
				$(id).removeClass("gif-load");
				if(data.StartTm != undefined) {
					console.log('StartTm=' + data.StartTm);
					$(id).find("td.StartTm span").text(fmtTm(data.StartTm));
					$(id).find("td.StartTm input").val(fmtTm(data.StartTm));
				}
				if(data.StartTm_i) {
					console.log('StartTm_i=' + data.StartTm_i);
					$(id).find("td.StartTm span").text(fmtTm(data.StartTm_i));
					$(id).find("td.StartTm input").val(fmtTm(data.StartTm_i));
					$(id).find("td.StartTm").addClass('modify');
				} else {
					$(id).find("td.StartTm").removeClass('modify');
				}
				if(data.FinishTm != undefined) {
					console.log('FinishTm=' + data.FinishTm);
					$(id).find("td.FinishTm span").text(fmtTm(data.FinishTm));
					$(id).find("td.FinishTm input").val(fmtTm(data.FinishTm));
				}
				if(data.FinishTm_i) {
					console.log('FinishTm_i=' + data.FinishTm_i);
					$(id).find("td.FinishTm span").text(fmtTm(data.FinishTm_i));
					$(id).find("td.FinishTm input").val(fmtTm(data.FinishTm_i));
					$(id).find("td.FinishTm").addClass('modify');
				} else {
					$(id).find("td.FinishTm").removeClass('modify');
				}
				if(data.Actual) {
					console.log('Actual=' + data.Actual);
					$(id).find("td.Actual span").text(fmtHour(parseFloat(data.Actual)));
					$(id).find("td.Actual input").val(fmtHour(parseFloat(data.Actual)));
				}
				if(data.Actual_i) {
					console.log('Actual_i=' + data.Actual_i);
					$(id).find("td.Actual span").text(fmtHour(parseFloat(data.Actual_i)));
					$(id).find("td.Actual input").val(fmtHour(parseFloat(data.Actual_i)));
					$(id).find("td.Actual").addClass('modify');
				} else {
					$(id).find("td.Actual").removeClass('modify');
				}
				if(data.Extra) {
					console.log('Extra=' + data.Extra + ' ' + fmtHour(parseFloat(data.Extra)));
					console.log($(id).find("td.Extra").html());
					console.log($(id).find("td.Extra span").attr('class'));
					console.log($(id).find("td.Extra input").attr('class'));
					$(id).find("td.Extra span").text(fmtHour(parseFloat(data.Extra)));
					$(id).find("td.Extra input").val(fmtHour(parseFloat(data.Extra)));
					console.log($(id).find("td.Extra").html());
				}
				if(data.Extra_i) {
					console.log('Extra_i=' + data.Extra + ' ' + fmtHour(parseFloat(data.Extra_i)));
					$(id).find("td.Extra span").text(fmtHour(parseFloat(data.Extra_i)));
					$(id).find("td.Extra input").val(fmtHour(parseFloat(data.Extra_i)));
					$(id).find("td.Extra").addClass('modify');
				} else {
					$(id).find("td.Extra").removeClass('modify');
				}
				if(data.Night) {
					console.log('Night=' + data.Night);
					$(id).find("td.Night span").text(fmtHour(parseFloat(data.Night)));
					$(id).find("td.Night input").val(fmtHour(parseFloat(data.Night)));
				}
				if(data.Night_i) {
					console.log('Night_i=' + data.Extra + ' ' + fmtHour(parseFloat(data.Night_i)));
					$(id).find("td.Night span").text(fmtHour(parseFloat(data.Night_i)));
					$(id).find("td.Night input").val(fmtHour(parseFloat(data.Night_i)));
					$(id).find("td.Night").addClass('modify');
				} else {
					$(id).find("td.Night").removeClass('modify');
				}
				if(data.Dayoff) {
					console.log('Dayoff=' + data.Dayoff);
					$(id).find("td.Dayoff span").text(fmtHour(parseFloat(data.Dayoff)));
					$(id).find("td.Dayoff input").val(fmtHour(parseFloat(data.Dayoff)));
				}
				if(data.Dayoff_i) {
					console.log('Dayoff_i=' + data.Extra + ' ' + fmtHour(parseFloat(data.Dayoff_i)));
					$(id).find("td.Dayoff span").text(fmtHour(parseFloat(data.Dayoff_i)));
					$(id).find("td.Dayoff input").val(fmtHour(parseFloat(data.Dayoff_i)));
					$(id).find("td.Dayoff").addClass('modify');
				} else {
					$(id).find("td.Dayoff").removeClass('modify');
				}
				total();
	        },
			onFail: function (jqXHR, textStatus, errorThrown) {
				console.log('onFail(jqXHR, textStatus, errorThrown)');
				console.log(jqXHR);
				console.log(textStatus);
				console.log(errorThrown);
				$('#msg').html(jqXHR.responseText);

			},
			onAlways: function () {
				console.log('onAlways()');
			},
		    columns: {
		      identifier: [0, 'id'],
		        editable: [  [ 5, 'Shift', shift_list]
							//,[ 6, 'BegTm_i']
							,[ 7, 'StartTm_i']
							,[ 8, 'FinishTm_i']
							//,[ 9, 'FinTm_i']
							,[10, 'Actual_i']
							,[11, 'Extra_i']
							,[12, 'Night_i']
							,[13, 'Dayoff_i']
							,[14, 'PTO']
							,[15, 'PTO_tm']
							,[16, 'Late']
							,[17, 'Early']
							,[18, 'Memo']
						  ]
		    }
		});
});
$(document).on("change",'.hour input[type="text"]', function() {
	console.log('change() this=' + this);
	console.log('change() val()=' + $(this).val());
	if($(this).val() != '0') {
		$(this).val(fmtHour(parseFloat($(this).val())));
	}
});
$(document).on("change",'.time input[type="text"]', function() {
	console.log('change() this=' + this);
	console.log('change() val()=' + $(this).val());
	console.log('change() parent()=' + $(this).parent().html());
	console.log('change() parent(title)=' + $(this).parent().attr('title'));
	var	tm = $(this).val().trim();
	if(tm == '') {
//		$(this).val($(this).parent().attr('title'));
	} else if(tm.match(/^\d{2}:\d{2}$/)) {
		// hh:mm
		console.log('hh:mm=' + tm);
	} else if(tm.match(/^\d{1}:\d{2}$/)) {
		// h:mm
		console.log('h:mm=' + tm);
		tm = '0' + tm;
		$(this).val(tm);
	} else if(tm.match(/^\d{4}$/)) {
		// hhmm
		console.log('hhmm=' + tm);
		tm = tm.slice(0,2) + ':' + tm.slice(2,4);
		$(this).val(tm);
	} else if(tm.match(/^\d{3}$/)) {
		// hmm
		console.log('hmm=' + tm);
		tm = '0' + tm.slice(0,1) + ':' + tm.slice(1,3);
		$(this).val(tm);
	} else {
		console.log('unmatch=' + tm);
		return false;
	}
});
$(document).ready(function() {
	//年月初期値セット
	var	dt = new Date();
	console.log('dt=' +dt);
	if (dt.getDate() > 18) {
		dt.setDate(1);
		dt.setMonth(dt.getMonth() + 1);
	}
	console.log('dt=' +dt);
//	dt.setDate(dt.getDate() + 10);
	var	month = dt.getFullYear() + '-' + ('0' + (dt.getMonth() + 1)).slice(-2);
	$("input[name='month']").val(month);
	window.location.hash = '#tab_month';
});
function total() {
	console.log('total()');
	var	id = '';
	var	days = 0;
	var	pto = 0;
	var	actual = 0;
	var	extra = 0;
	var	night = 0;
	var	dayoff = 0;
	var	slist = 0;
	$('#slist').find("option").remove();
	$('#list tbody tr').each(function(i) {
//		console.log(i + ' ' + id + ' ' + this.id);
		if(id != this.id.slice(0,5)) {
			id = this.id.slice(0,5);
			days = 0;
			pto = 0;
			actual = 0;
			extra = 0;
			night = 0;
			dayoff = 0;
		    $('#slist').append($('<option />').val(id).html($(this).find(".Name").text()));
			slist++;
		}
		if($(this).find(".StartTm").text() || $(this).find(".PTO").text() || $(this).find(".PTO_tm").text() ) {
			days++;
		}
		$(this).find(".Days").text(days);

		if($(this).find(".PTO").text()) {
			pto += parseFloat($(this).find(".PTO").text());
		}
		if($(this).find(".PTO_tm").text()) {
			pto += parseFloat($(this).find(".PTO_tm").text());
		}
		$(this).find(".PTO_H").text(fmtHour(pto));

		if($(this).find(".Actual").text()) {
			actual += parseFloat($(this).find(".Actual").text());
		}
		$(this).find(".Actual_H").text(fmtHour(actual));

		if($(this).find(".Extra").text()) {
			extra += parseFloat($(this).find(".Extra").text());
		}
		$(this).find(".Extra_H").text(fmtHour(extra));

		if($(this).find(".Night").text()) {
			night += parseFloat($(this).find(".Night").text());
		}
		$(this).find(".Night_H").text(fmtHour(night));

		if($(this).find(".Dayoff").text()) {
			dayoff += parseFloat($(this).find(".Dayoff").text());
		}
		$(this).find(".Dayoff_H").text(fmtHour(dayoff));

		$(this).find(".Total_H").text(fmtHour(actual + extra + night + dayoff + pto));
	});
}
var	thead = '';
$('input[name="copy0"]').on('click', function(){
	console.log(this);
	thead = $('#list thead').html();
	$.when(
		$('#list th').each(function(i) {
			console.log($(this).html());
			$(this).html($(this).html().replace('<br>', ''));
		})
	).done(function() {
		$('#tab_month input[name="copy"]').trigger('click');	//編集可
	});
});
$('input[name="copy"]').on('click', function(){
	console.log(this);
    var clipboard = new ClipboardJS(this);
    clipboard.on('success', function(e) {
		console.log(e);
		$('#list thead').html(thead);
		clipboard.destroy();
    });
    clipboard.on('error', function(e) {
		console.log(e);
		$('#list thead').html(thead);
		clipboard.destroy();
    });
});
//従業員テーブル 検索
$("#tab_staff input[name='staff']").on('click', function() {
	var	req = 'staff.py?dns=' + $('#dns').val();
	if($("#quit").prop("checked")) {
		req += '&quit=true';
	}
	$('#msg').text(req);
	$(this).addClass("gif-load");
	fetch(req).then((res) => {
		$('#msg').text('');
		$(this).removeClass("gif-load");
		return res.json();
	}).then((json) => {
		var	tr = '';
		for ( var i = 0 ; i < json.data.length ; i++) {
//			for (var item in json.data[i]) {
//				tr += '<td>' + json.data[i][item] || '' + '</td>';
//			}
			tr += '<tr id="' + json.data[i].StaffNo + '">';
			tr += '<td>' + json.data[i].StaffNo + '</td>';
			tr += '<td>' + json.data[i].Post + '</td>';
			tr += '<td>' + json.data[i].StaffNo + '</td>';
			tr += '<td>' + json.data[i].Name + '</td>';
			tr += '<td>' + json.data[i].Shift + '</td>';
			tr += '<td>' + fmtDt(json.data[i].QuitDt) + '</td>';
			tr += '<td>' + json.data[i].Quit + '</td>';
			tr += '</tr>';
		}
		$('#tab_staff table tbody').find("tr").remove();
		$('#tab_staff table tbody').append(tr);
		$('#tab_staff table').Tabledit({
			url: 'staff.py?dns=' + $('#dns').val(),
			editButton: false,
			deleteButton: false,
			hideIdentifier: true,
		    columns: {
		      identifier: [0, 'id'],
		        editable: [  [ 1, 'Post']
							,[ 2, 'StaffNo']
							,[ 3, 'Name']
							,[ 4, 'Shift']
							,[ 5, 'QuitDt']
							,[ 6, 'Quit']
						  ]
		    },
			onDraw: function () {
				console.log('onDraw()');
//				console.log($('#tab_staff table tbody').html());
			},
			onAjax: function(action, serialize) {
				//Ajax開始時
				console.log('onAjax(action, serialize)');
				console.log(action);
				console.log(serialize);
				var urlParams = new URLSearchParams(serialize);
				console.log(urlParams.get('id'));
			},
			onSuccess: function(data, textStatus, jqXHR) {
				console.log('onSuccess(data, textStatus, jqXHR)');
				console.log(data);
				console.log(textStatus);
				console.log(jqXHR);
				console.log(data.id);
	        },
			onFail: function (jqXHR, textStatus, errorThrown) {
				console.log('onFail(jqXHR, textStatus, errorThrown)');
				console.log(jqXHR);
				console.log(textStatus);
				console.log(errorThrown);
			},
			onAlways: function () {
				console.log('onAlways()');
			},
		});
	}).catch((error) => {
		$('#tab_staff table tbody').find("tr").remove();
		$('#tab_staff table tbody').append('<tr><td>' + error + '</td></tr>');
	});
});
$('#tab_staff').trigger('click');
/* ------------------------------
 Loading イメージ表示関数
 引数： msg 画面に表示する文言
 ------------------------------ */
function dispLoading(msg){
  // 引数なし（メッセージなし）を許容
  if( msg == undefined ){
    msg = "";
  }
  // 画面表示メッセージ
  var dispMsg = "<div class='loadingMsg'>" + msg + "</div>";
  // ローディング画像が表示されていない場合のみ出力
  if($("#loading").length == 0){
    $("body").append("<div id='loading'>" + dispMsg + "</div>");
  }
}
 
/* ------------------------------
 Loading イメージ削除関数
 ------------------------------ */
function removeLoading(){
  $("#loading").remove();
}
/* ------------------------------
勤務表 印刷
 ------------------------------ */
$("#tab_month input[name='print']").on('click', function() {
	var	title = '勤務表-' + $("input[name='month']").val();;
	var	html = '<html><head><title>' + title + '</title>';
	html += '</head>';
	html += '<style>';
	html += 'table {';
	html += 'border-collapse: collapse;';
	html += '}';
	html += 'table caption {';
	html += 'text-align : left;';
	html += 'font-size : 2em;';
	html += '}';
	html += 'th {';
	html += 'padding:1px 1px;';
	html += 'font-weight: normal;';
	html += 'border: solid thin black;';
	html += 'white-space: nowrap;';
	html += '}';
	html += 'td {';
	html += 'padding:6px 5px;';
	html += 'font-weight: normal;';
	html += 'border: solid thin black;';
	html += 'white-space: nowrap;';
	html += '}';
	html += 'tr th:nth-child(-n+4) {display : none;}';
	html += 'tr td:nth-child(-n+4) {display : none;}';
	html += 'tr td:nth-child(n+5):nth-child(-n+6) {text-align : center;}';
	html += 'tr th:nth-child(6) {font-size:0.6em}';
	html += 'tr td:nth-child(7),td:nth-child(10) {font-size:0.6em;}';
	html += 'tr td:nth-child(n+11):nth-child(-n+16) {text-align : right;}';
	html += 'tr td:nth-last-child(8) {white-space: normal; font-size:0.6em;padding:1;}';	//備考
	html += 'tr td:nth-last-child(-n+7) {text-align : right;}';
//	html += 'tr.page-break {background: yellow;}';
//	html += '@media print {';
//	html += 'table {display: table;}';
//	html += 'td {display: table-cell;}';
//	html += 'tr.page-break  { display: block; page-break-before: always; }';
//	html += '}';
	html += '</style>';
	html += '<body onLoad="window.print();">';
	html += '<table>';
	var	thead = '<thead>';
	$('#list thead tr').each(function(i) {
		thead += '<tr>';
		$(this).find('th').each(function(j) {
			thead += '<th>' + $(this).html() + '</th>';
		});
		thead += '</tr>';
	});
	thead += '</head>';
	$('#list tbody tr').each(function(i) {
		if(i > 0 && $(this).hasClass('dt-top')) {
			html += '</tbody></table>';
			html += '<table style="page-break-before:always;">';
		} else if(i == 0) {
			html += '<table>';
		}
		if(i == 0 || $(this).hasClass('dt-top')) {
			html += '<caption>';
			var	month = $("input[name='month']").val();
			html += '<span>' + month.replace('-0','-').replace('-','年') + '月</span>';
			html += ' <span>' + $(this).find('.Post').text() + '</span>';
			html += ' <span>' + $(this).find('.StaffNo').text() + '</span>';
			html += ' <span>' + $(this).find('.Name').text() + '</span>';
			html += '</caption>';
			html += thead;
			html += '<tbody>';
		}
		html += '<tr>';
		$(this).find('td').each(function(j) {
			html += '<td>';
			if($(this).find('span').text() != '') {
				html += $(this).find('span').text();
			} else {
				html += $(this).text();
			}
			html += '</td>';
		});
		html += '</tr>';
	});
	html += '</tbody>';
	html += '</table>';
	html += '</body></html>';
	var	w = window.open("", title);
	w.document.write(html);
	w.document.close();
});
$("input[name='staff_add']").on('click', function() {
    var table = $(this).attr('for-table');  //get the target table selector
    var $tr = $(table + ">tbody>tr:last-child").clone(true, true);  //clone the last row
    var nextID = parseInt($tr.find("input.tabledit-identifier").val()) + 1; //get the ID and add one.
    $tr.find("input.tabledit-identifier").val(nextID);  //set the row identifier
    $tr.find("span.tabledit-identifier").text(nextID);  //set the row identifier
    $(table + ">tbody").append($tr);    //add the row to the table
    $tr.find(".tabledit-edit-button").click();  //pretend to click the edit button
    $tr.find("input:not([type=hidden]), select").val("");   //wipe out the inputs.
});
