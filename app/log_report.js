/*
log_report.js
*/
$(document).ready(function() {
	// バージョン
	var	ver = [];
	ver.push('0.00 2021.07.28 sebd作業ログ集計');
	$('#msg').html(ver.slice(0,1).join('<br>'));
	$('#version').html(ver.join('<p>'));
	// 出荷日
	var today = new Date();
	var dt2 = today.getFullYear()+("0"+(today.getMonth()+1)).slice(-2)+( "0"+today.getDate()).slice(-2);
	if ( today.getDate() < 21 ) {
		today.setMonth(today.getMonth() - 1);
	}
	var dt1 = today.getFullYear()+
		( "0"+( today.getMonth()+1 ) ).slice(-2)+
		"21";
	$('#date_st').val(dt1);
	$('#date_ed').val(dt2);

	// 検索ボタン
	$('#submit').on('click', function() {
		console.log('■検索開始.' + this.id);
		var	req = 'log_report.py?dns=' + $('#dns').val();
		req += '&s=' + $('#date_st').val();
		req += '&e=' + $('#date_ed').val();
		console.log(req);
		$('#log').html($('#log').html() + '<br>' + req);
		$(this).addClass("gif-load");
		fetch(req)
		.then((res) => {
			$(this).removeClass("gif-load");
			return res.json();
		})
		.then((json) => {
			$('#list').find("tr").remove();
			var	tr = '';
			tr = '<tr>';
			tr += '<th rowspan="2" class="row"></th>';
			tr += '<th rowspan="2">工程/要因</th>';
			tr += '<th colspan="2">合計</th>';
			for (var item in json.columns) {
				if(json.columns[item][0] == 'Cnt') {
					var	md = json.columns[item][1];
					md = parseInt(md.slice(4,6)) + '/' + parseInt(md.slice(6));
					tr += '<th colspan="2" title="' + json.columns[item][1] + '">' + md + '</th>';
				}
			}
			tr += '</tr>';
			tr += '<tr>';
			tr += '<th>件</th>';
			tr += '<th>個</th>';
			for (var item in json.columns) {
				if(json.columns[item][0] == 'Cnt') {
					tr += '<th>件</th>';
					tr += '<th>個</th>';
				}
			}
			tr += '</tr>';
			$('#list thead').append(tr);
			tr = '';
			var	kotei = '';
			for ( var i = 0 ; i < json.index.length ; i++ ) {
				if(kotei != json.index[i][0]) {
					kotei = json.index[i][0];
					console.log(kotei);
					tr += '<tr>';
					tr += '<td class="row"></td>';
					//工程
					tr += '<td class="' + kotei + '">'+ kotei + '</td>';
					//合計
					var	cnt = 0;
					var	qty = 0;
					for ( var sum_i = 0 ; sum_i < json.index.length ; sum_i++ ) {
						if(kotei == json.index[sum_i][0]) {
							for ( var sum_j = 0 ; sum_j < json.data[sum_i].length ; sum_j++ ) {
								if(sum_j < (json.data[sum_i].length / 2)) {
									cnt += json.data[sum_i][sum_j];
								} else {
									qty += json.data[sum_i][sum_j];
								}
							}
						}
					}
					tr += '<td class="number">' + cnt + '</td><td class="number">' + qty + '</td>';
					//日別
					var	k = json.data[i].length / 2;
					for ( var j = 0 ; j < k ; j++ ) {
						var	cnt = 0;
						var	qty = 0;
						for ( var sum_i = 0 ; sum_i < json.index.length ; sum_i++ ) {
							if(kotei == json.index[sum_i][0]) {
								cnt += json.data[sum_i][j];
								qty += json.data[sum_i][j + k];
							}
						}
						tr += '<td class="number">' + cnt + '</td>';
						tr += '<td class="number">' + qty + '</td>';
					}
					tr += '</tr>';
				}
				tr += '<tr class="yoin">';
				//#
				tr += '<td class="row"></td>';
				//工程/要因
				tr += '<td class="yoin ' + json.index[i][0] + '">'+ json.index[i][1] + '</td>';
				//合計
				var	cnt = 0;
				var	qty = 0;
				for ( var j = 0 ; j < json.data[i].length ; j++ ) {
					if(j < (json.data[i].length / 2)) {
						cnt += json.data[i][j];
					} else {
						qty += json.data[i][j];
					}
				}
				tr += '<td class="number">' + cnt + '</td><td class="number">' + qty + '</td>';
				//日別
				var	k = json.data[i].length / 2;
				for ( var j = 0 ; j < k ; j++ ) {
					tr += '<td class="number">' + (json.data[i][j] || '') + '</td>';
					tr += '<td class="number">' + (json.data[i][j + k] || '') + '</td>';
				}
				tr += '</tr>';
			}
			$('#list tbody').append(tr);
		}).catch((err) => {
			console.log('catch:' + err);
			$('#msg').html(req + '<p>' + err);
		});
		return false;
	});
	$('#copy').on('click', function(){
	    var clipboard = new Clipboard('#copy');
	    clipboard.on('success', function(e) {
			//成功時の処理
	    });
	    clipboard.on('error', function(e) {
			//失敗時の処理
	    });
	});
});
