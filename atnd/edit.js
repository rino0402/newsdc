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

//テーブル 検索
$("input[name='load']").on('click', function() {
	var	req = 'load.py?dns=' + $('#dns').val();
	req += '&month=' + $("input[name='month']").val();

	$(this).addClass("gif-load");
	$('#msg').text(req);
	fetch(req).then((res) => {
		$('#msg').text('');
		$(this).removeClass("gif-load");
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
			tr += '<tr id="' + id + '">';
			tr += '<td>' + id + '</td>';
			tr += '<td>' + json.data[i]['StaffNo'] + '</td>';
			tr += '<td>' + json.data[i]['Name'] + '</td>';
			var	cls = '';
			if (json.data[i]['Holiday'] != '') {
				cls = ' holiday';
			}
			tr += '<td class="date ' + json.data[i]['strDay'] + cls + '" title="' + json.data[i]['Holiday'] + '">' + json.data[i]['fmtDt'] + '</td>';
//			tr += '<td>' + '</td>';
//			tr += '<td>' + json.data[i]['Shift'] + '</td>';
			if (json.data[i]['BegTm_i'] == '') {
				tr += '<td class="time BegTm">' + json.data[i]['BegTm5'] + '</td>';
			} else {
				tr += '<td class="time BegTm modify" title="' + json.data[i]['BegTm5'] + '">';
				tr += json.data[i]['BegTm_i'] + '</td>';
			}
			if (json.data[i]['FinTm_i'] == '') {
				tr += '<td class="time FinTm">' + json.data[i]['FinTm5'] + '</td>';
			} else {
				tr += '<td class="time FinTm modify" title="' + json.data[i]['FinTm5'] + '">';
				tr += json.data[i]['FinTm_i'] + '</td>';
			}
			tr += '<td class="hour">' + fmtHour(json.data[i]['Late']) + '</td>';
			tr += '<td class="hour">' + fmtHour(json.data[i]['Early']) + '</td>';
			tr += '<td class="hour PTO">' + fmtHour(json.data[i]['PTO']) + '</td>';
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
			tr += '<td class="Memo">' + json.data[i]['Memo'] + '</td>';
			tr += '<td class="total Days"></td>';
			tr += '<td class="total PTO_H"></td>';
			tr += '<td class="total Actual_H"></td>';
			tr += '<td class="total Extra_H"></td>';
			tr += '<td class="total Night_H"></td>';
			tr += '<td class="total Total_H"></td>';
//			for (var item in json.data[i]) {
//				tr += '<td>' + json.data[i][item] || '' + '</td>';
//			}
			tr += '</tr>';
		}
		$('#list tbody').find("tr").remove();
		$('#list tbody').append(tr);
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
				if(data.BegTmM !== 'undefined') {
					if(data.BegTmM == '') {
//						$(id).find(".BegTm").removeClass('modify');
						$(id).find(".BegTm").text($(id).find(".BegTm").attr('title'));
					} else {
//						$(id).find(".BegTm").addClass('modify');
					}
				}
				if(data.Actual) {
					console.log('Actual=' + data.Actual);
					$(id).find("td.Actual span").text(fmtHour(parseFloat(data.Actual)));
					$(id).find("td.Actual input").val(fmtHour(parseFloat(data.Actual)));
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
				if(data.Night) {
					console.log('Night=' + data.Night);
					$(id).find("td.Night span").text(fmtHour(parseFloat(data.Night)));
					$(id).find("td.Night input").val(fmtHour(parseFloat(data.Night)));
				}
				total();
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
		    columns: {
		      identifier: [0, 'id'],
		        editable: [	 [ 4, 'BegTm_i']
							,[ 5, 'FinTm_i']
							,[ 6, 'Late']
							,[ 7, 'Early']
							,[ 8, 'PTO']
							,[ 9, 'Actual_i']
							,[10, 'Extra_i']
							,[11, 'Night_i']
							,[12, 'Memo']
						  ]
		    }
		});
		$('#list').exTableFilter('#filter');
		total();
	}).catch((error) => {
		$('#msg').text(error);
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
		$(this).val($(this).parent().attr('title'));
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
	var	dt = new Date();
	dt.setDate(dt.getDate() + 10);
	var	month = dt.getFullYear() + '-' + (dt.getMonth() + 1);
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
	$('#list tbody tr').each(function(i) {
//		console.log(i + ' ' + id + ' ' + this.id);
		if(id != this.id.slice(0,5)) {
			id = this.id.slice(0,5);
			days = 0;
			pto = 0;
			actual = 0;
			extra = 0;
			night = 0;
		}
		if($(this).find(".BegTm").text()) {
			days++;
		}
		$(this).find(".Days").text(days);

		if($(this).find(".PTO").text()) {
			pto += parseFloat($(this).find(".PTO").text());
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
		$(this).find(".Total_H").text(fmtHour(actual + extra + night));
	});
}
$('input[name="copy"]').on('click', function(){
	console.log(this);
	$('#list th').each(function(i) {
		console.log($(this).html());
		$(this).data('html', $(this).html());
		$(this).html($(this).html().replace('<br>', ''));
	});
    var clipboard = new ClipboardJS(this);
    clipboard.on('success', function(e) {
		console.log(e);
		$('#list th').each(function(i) {
			console.log($(this).data('html'));
			$(this).html($(this).data('html'));
		});
    });
    clipboard.on('error', function(e) {
		console.log(e);
    });
});

$(function () {
	$('--#list').Tabledit({
		url: 'edit.py',
		columns: {
			identifier: [0, 'id'],
			editable: [[6, 'bgnTm'], [7, 'finTm'], [8, 'Actual']]
	    },
		onDraw: function () {
			console.log('onDraw()');
		},
        onAjax: function(action, serialize) {
            console.log('onAjax(action, serialize)');
            console.log(action);
            console.log(serialize);
        },
        onSuccess: function(data, textStatus, jqXHR) {
            console.log('onSuccess(data, textStatus, jqXHR)');
            console.log(data);
            console.log(textStatus);
            console.log(jqXHR);
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
        buttons: {
            edit: {
                class: 'btn btn-sm btn-default',
                html: '<span class="glyphicon glyphicon-pencil"></span>',
                action: 'edit'
            },
            delete: {
                class: 'btn btn-sm btn-default',
                html: '<span class="glyphicon glyphicon-trash"></span>',
                action: 'delete'
            },
            save: {
                class: 'btn btn-sm btn-success',
                html: '保存'
            },
            restore: {
                class: 'btn btn-sm btn-warning',
                html: 'Restore',
                action: 'restore'
            },
            confirm: {
                class: 'btn btn-sm btn-danger',
                html: '削除'
            }
		}
	});
});