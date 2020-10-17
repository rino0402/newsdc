/*
monitor.js
https://material.io/resources/icons/?style=baseline
2017.12.12 海外出荷業務メッセージ表示
*/
//	version = '0.11 2018.01.09 起動時にバージョン表示';
//	version = '0.12 2018.01.09 海外出荷注意事項のスライド表示';
//	version = '0.13 2018.01.10 海外出荷注意事項のスライド表示';
//	version = '0.14 2018.01.10 スライド表示の時にチャイム音';
//var	version = '0.15 2018.01.11 Activeデータ受信でチャイム音';
//var	version = '0.16 2018.01.23 予定と完了を在庫データから算出';
//var	version = '0.17 2018.01.26 他センター対応:滋賀pc/dc';
//var	version = '0.18 2018.02.28 商品化済を非表示';
//var	version = '0.19 2018.03.13 monitor.js';
//var	version = '0.20 2018.03.16 大阪対応';
//var	version = '0.21 2018.03.17 大阪対応';
//var	version = '0.22 2018.03.18 y_syuka対応';
//var	version = '0.23 2018.03.19 y_syuka対応';
//var	version = '0.24 2018.03.19 p_sshiji対応';
//var	version = '0.25 2018.03.21 スクロール対応';
//var	version = '0.26 2018.03.22 nav対応';
//var	version = '0.26 2018.03.23 画面切替不具合修正';
//var	version = '0.27 2018.03.24 css整理';
//var	version = '0.28 2018.03.25 作業状況';
//var	version = '0.29 2018.03.27 余計なfetchをしないように修正';
//var	version = '0.30 2018.03.28 ';
//var	version = '0.31 2018.03.29 ';
//var	version = '0.32 2018.04.02 ';
//var	version = '0.33 2018.04.03 EQこころの鍛え方カレンダー';
//var	version = '0.34 2018.04.03 矢印キー操作中に画面自動切替しないように変更';
//var	version = '0.35 2018.04.05 メッセージ表示の整理';
//var	version = '0.36 2018.04.06 メッセージ表示の整理';
//var	version = '0.36 2018.04.07 エアコン発注残';
//var	version = '0.37 2018.04.08 エアコン発注残';
//var	version = '0.38 2018.04.09 エアコン発注残 表示内容変更';
//var	version = '0.39 2018.04.09 入庫待ち在庫';
//var	version = '0.40 2018.04.10 メッセージの色変更';
//var	version = '0.41 2018.04.10 エアコン発注残 指定納期を全て月/日で表示';
//var	version = '0.42 2018.04.10 資材入荷予定 収支指定：A=120 R=220';
//var	version = '0.43 2018.04.12 スクロール中はページ切替しない';
//var	version = '0.44 2018.04.13 出荷状況（運送会社別） キャンセル明細表示';
//var	version = '0.45 2018.04.20 商品化状況：商品化計画データを表示';
//var	version = '0.46 2018.04.20 商品化状況：エアコン対応';
//var	version = '0.47 2018.05.16 商品化状況：エアコン当日分のみ表示';
//var	version = '0.48 2018.06.06 商品化状況：計画データ参照不具合修正';
//var	version = '0.49 2018.06.20 海外出荷状況：規制注意対応中 & 見やすく変更中...';
//var	version = '0.50 2018.06.20 海外出荷：注意事項対応(滋賀PC)';
//var	version = '0.51 2018.06.21 海外出荷：注意事項対応(滋賀PC/小野PC)';
//var	version = '0.52 2018.06.22 出荷状況（運送会社別） 未出庫明細表示';
var	version = '<p>0.53 2018.06.25 出荷状況（運送会社別） 未出庫明細表示 ※未出庫10件未満で表示するように変更';
var	version = '<p>0.54 2018.06.27 商品化時の注意事項を「原産国・品番・色・形状」に変更';
var	version = '<p>0.55 2018.07.13 出荷状況（小野,滋P）：「送信残」追加';
var	version = '<p>0.56 2018.08.24 出荷状況（小野,滋P）：前日未出荷分を表示';
var	version = '<p>0.57 2018.08.29 海外出荷状況（小野,滋P）：先行出荷対応';
var	version = '<p>0.58 2018.08.30 海外出荷状況（小野,滋P）：出荷日変更対応';
var	version = '<p>0.59 2018.09.03 サーバー負荷低減の為、同時処理しないように変更';
var	version = '<p>0.60 2018.09.06 画面の更新時刻を右上に表示';
var	version = '<p>0.61 2018.09.08 画面の更新時刻を右上に表示 y_syuka order';
var	version = '<p>0.62 2018.09.13 処理順変更 y_syuka_h y_syuka order p_sshiji';
var	version = '<p>0.62 2018.09.13 出荷状況（小野,滋P）：出荷先名はみ出しカット';
var	version = '<p>0.62 2018.09.13 出荷状況（小野,滋P）：3サテ背景色対応';
var	version = '<p>0.63 2018.09.14 大阪事：入庫残件数表示';
var	version = '<p>0.64 2018.09.18 Active出荷データ受信 チャイム音変更（小野,滋P）';
var	version = '<p>0.65 2018.09.19 出荷状況（小野,滋P）：出庫／出荷残を見やすく(黄,黒)しました...ﾂﾓﾘ';
var	version = '<p>0.66 2018.09.20 欠品状況（小野,滋P）を追加しました.';
var	version = '<p>0.67 2018.09.27 出荷状況（大阪）便数の表示形式変更／伝票枚数タイトルの背景色変更.';
var	version = '<p>0.68 2018.09.28 出荷状況（大阪）便数の表示形式変更 コロン(:)を除外';
var	version = '<p>0.69 2018.10.01 海外出荷注意事項（小野,滋P）：表示が崩れるのを修正／表示タイミング変更';
var	version = '<p>0.70 2018.10.02 先頭行が固定されない不具合修正';
var	version = '<p>0.71 2018.10.16 欠品状況（小野,滋P）：国内/海外別に集計.';;
var	version = '<p>0.72 2018.10.17 音声出力対応中...';;
var	version = '<p>0.73 2018.10.24 <span class="h2">商品化状況：指図票を発行していない計画を非表示／状態にチェック済を表示</span>';
var	version = '<p>0.74 2018.10.26 <span class="h2">商品化状況：チェック済をチェックマーク<span class="check icon"></span>　表示</span>';
var	version = '<p>0.75 2018.11.01 表示ページが切り替わらない不具合を修正';
var	version = '<p>0.76 2018.12.28 出荷状況（滋P）：合計欄を件数のみに変更';
var	version = '<p>0.77 2019.02.04 出荷状況（滋P）：列の並びを補充／緊急に変更';
var	version = '<p>0.78 2019.02.05 出荷状況（滋P）：引当済(進捗3)を表示';
var	version = '<p>0.79 2019.02.06 出荷商品化待ち：【新規対応】';
var	version = '<p>0.80 2019.02.06 出荷商品化待ち：小野対応';
var	version = '<p>0.81 2019.02.07 出荷商品化待ち：92在庫数を追加';
var	version = '<p>0.82 2019.02.08 出荷商品化待ち：倉庫(標準棚倉庫名)を追加';
var	version = '<p>0.83 2019.02.12 出荷状況（小野,滋P）：個数,才数を表示';
var	version = '<p>0.84 2019.02.13 出荷状況（滋P）：合梱(AMV69Z-JS0 AMV70Z-LS0)の才数に対応';
var	version = '<p>0.85 2019.02.23 出荷商品化待ち：品番ごとの合計で集計するように変更';
var	version = '<p>0.85 2019.03.01 Merry Hina Party!';
var	version = '<p>0.85 2019.03.05 Saboten♪';
var	version = '<p>0.86 2019.03.15 出荷商品化待ち：滋賀PC 倉庫カラー表示';
var	version = '<p>0.87 2019.03.18 morning service♪';
var	version = '<p>0.88 2019.03.22 出荷商品化待ち：表示設定対応';
var	version = '<p>0.89 2019.03.22 商品化状況：(滋賀PC)出荷ありを上位に表示';
var	version = '<p>0.90 2019.03.27 商品化状況：欠品数を表示';
var	version = '<p>0.91 2019.04.05 2019年度スローガン表示';
var	version = '<p>0.92 2019.04.16 出荷状況（小野,滋P）：1画面に2日分表示できるように変更';
var	version = '<p>0.99 2019.06.26 出荷状況（小野,滋P）：当日の出庫残／検品残／Active送信残＞０で反転表示';
var	version = '<p>1.00 2019.07.02 欠品状況：エアコン追加(緊急欠品のみ)';
var	version = '<p>1.01 2019.08.16 商品化状況：(滋賀PC)出荷数＞商品化済在庫＆92在庫＞0を上位に表示';
var	version = '<p>1.02 2019.09.23 同時処理によりサーバーが高負荷になるのを防止';
var	version = '<p>1.03 2019.10.04 ※海外出荷《注意事項》：<br>行間を縮小 10px→1px／表示時間を短縮 30秒→15秒／滋賀PC出荷用で午後から表示しない';
var	version = '<p>1.04 2019.10.23 誤出荷注意メッセージ追加';
var	version = '<p>1.05 2019.11.06 広島事 初期設定対応';
var	version = '<p>1.06 2019.11.07 広島事 JCS納品リスト';
var	version = '<p>1.07 2019.11.14 一部の画面が切り替わらない不具合修正';
var	version = '<p>1.08 2019.11.25 ワイヤレスポインターの右矢印ボタンで次画面に切り替わるように対応';
var	version = '<p>1.09 2019.11.25 ワイヤレスポインターの右ボタンで次画面に切り替わるように対応';
var	version = '<p>1.10 2019.11.26 スクロール時のヘッダー固定位置を画面最上部に変更';
var	version = '<p>1.11 2019.12.01 資材補充アラーム：新規対応';
var	version = '<p>1.12 2019.12.06 リモコン操作の不具合修正';
var	version = '<p>1.13 2019.12.11 JCSオーダー(区分別) 集計条件変更';
var	version = '<p>1.14 2019.12.12 JCSオーダー(区分別) Oリング 追加';
var	version = '<p>1.15 2019.12.13 JCSオーダー(区分別) パイプ(商品化タイプ:10S,10M,10L) 追加';
var	version = '<p>1.16 2019.12.16 JCSオーダー(区分別) NGを集計しないように修正';
var	version = '<p>1.17 2019.12.26 JCSオーダー概況 NGしかない納入日をスキップ';
var	version = '<p>1.18 2019.12.27 本年も大変お世話になりありがとうございました。よいお年を。';
var	version = '<p>1.19 2020.01.05 令和２年 今年もよろしくお願いいたします。';
var	version = '<p>1.20 2020.01.21 誤出荷注意メッセージ：2019年度誤出荷の内訳を表示';
var	version = '<p>1.21 2020.02.06 JCSオーダー(区分別) NG=Z(在庫なし)を集計';
var	version = '<p>1.22 2020.02.25 袋井 商品化予定 対応中...';
var	version = '<p>1.23 2020.02.27 滋賀pc 商品化予定 対応中...';
var	version = '<p>1.24 2020.03.07 商品化予定 商品化完了したものを下に表示';
var	version = '<p>1.25 2020.03.09 商品化予定 数量＝空白の場合は品番以降の項目を結合して表示';
var	version = '<p>1.26 2020.04.01 2020年度 スローガン「衆知を集め 全員経営」';
var	version = '<p>1.27 2020.05.27 広島事 出勤予定 行間を狭くして表示行数を増やしました.';
var	version = '<p>1.28 2020.06.05 誤出荷注意メッセージ：2020年度.';
var	version = '<p>1.29 2020.08.11 袋井「入出荷予定」対応しました.';
var	version = '<p>1.30 2020.08.25 大阪「出勤予定」テストバージョン.5';
var	version = '<p>1.31 2020.08.28 JCSオーダー概況 進捗チェックの不具合修正';
var	version = '<p>1.32 2020.10.05 JCSオーダー概況 担当者IDが空白になる不具合により進捗④にならないのを修正';
var	version = '<p>1.33 2020.10.06 袋井「欠品入荷リスト」対応しました.';
/*
debugフラグ
*/
var	debug_flag = false;
function debug(flag) {
	if(flag == true) {
		debug_flag = true;
	} else if(debug_flag == true) {
		$.toast({text: flag,loader: false});
	}
	return debug_flag;
}
function debug_toast(t) {
	var debug_to = $('#debug').val();
	if(debug_to > 0) {
		if(t.slice(0,1) != '<') {
			t = '<div class="h7">' + t + '</div>';
		}
		$.toast({
			text : t
			,loader: false
			,hideAfter : debug_to * 1000
			,stack : 8
		});
	}
}
/*
数値フォーマット
*/
var nFormat = function(number,n) {
	var	result = '';
	if (number != 0) {
		result = parseFloat(number).toFixed(n);
	}
    return result;
};
//設定：読込セット
function setConfig(id,def) {
	var v = localStorage.getItem($('#pref').text() + id);
	debug_toast( 'getItem() ' + $('#pref').text() + id + ':' + v );
	if(v == null) {
		v = def;
	}
	$(id).val(v);
	return v;
}
/*
メイン
*/
$(document).ready(function() {
	console.log('$(document).ready():' + this.title);
	$('#pref').text(location.search.substring(1));
	//バージョン表示
	$('#version').html(	version
						+ '<br>JQuery : '	+ $.fn.jquery
						);
	setConfig('#width',window.innerWidth);
	setConfig('#height',window.innerHeight);
	$('#SetSize').on('click', function(){
//		window.moveTo(0,0);
//		window.resizeTo($('#width').val(), $('#height').val());
/*
		$.toast({
			text : 'location.href：' + location.href
//			text : 'サイズ変更：' + $('#width').val() + ' x ' + $('#height').val()
			,showHideTransition : 'slide'	// 表示・消去時の演出
			,allowToastClose : true			// 閉じるボタンの表示・非表示
			,hideAfter : 10000				// 自動的に消去されるまでの時間(ミリ秒)
			,loader: false
		});
*/
		window.open(location.href,'_blank', 'width=' + $('#width').val() +', height=' + $('#height').val());
		return false;
	});
	var	stat = 1;
	//設定：変更メソッド
	$('.config').change(function() {
		var id = $('#pref').text() + '#' + this.id;
		console.log( 'change():' + id + ':' + $(this).val());
		try {
			localStorage.setItem(id,$(this).val());
			debug_toast('setItem ' + id + ':' + $(this).val());
		} catch(e) {}
	});
	//設定：初期値
	setConfig('#dns','newsdc');
	setConfig('#JGYOBU','');
	setConfig('#Soko','');
	setConfig('#debug',0);
	setConfig('#IntervalNotice',15);
	setConfig('#UKEHARAI_CODE','');
	setConfig('#sumi_disp','false');
	title = setConfig('#TitleName',$('#pref').text());
	if(title == '') {
		title = $('#dns').val();
	}
	//タイトルセット
	$('title').text(title);
	$('#name').text(title);
	setConfig('#volume','');
	//スクロールバー表示／非表示
	if( setConfig('#ScrollBar','true') == 'true') {
		$('body').css('overflow-y','auto');
	}
	// スライド表示
	var slide_dir = setConfig('#slide','');
	var slide_int = setConfig('#slide_interval','180');
	var slide_dly = setConfig('#slide_delay','30');
	if (slide_dir != '') {
		debug_toast('slide_dir:' + slide_dir
				  + '<br><br>slide_interval:' + slide_int
				  + '<br><br>slide_delay:' + slide_dly
					);
		setInterval(function(){
			$('.centerWindow').text('');
			$('.centerWindow').addClass('slide02');
			var	url = slide_dir + '/slide02.jpg';
			var now = new Date();
			if (now.getSeconds() <= 10) {
				url = slide_dir + '/slide02.jpg';
			} else if (now.getSeconds() <= 20) {
				url = slide_dir + '/slide03.jpg';
			} else if (now.getSeconds() <= 30) {
				url = slide_dir + '/slide04.jpg';
			} else if (now.getSeconds() <= 35) {
				url = slide_dir + '/slide05.jpg';
			} else if (now.getSeconds() <= 40) {
				url = slide_dir + '/slide06.jpg';
			} else if (now.getSeconds() <= 50) {
				url = slide_dir + '/slide07.jpg';
			} else if (now.getSeconds() <= 55) {
				url = slide_dir + '/slide08.jpg';
			} else if (now.getSeconds() <= 60) {
				url = slide_dir + '/slide09.jpg';
			}
			var xhr;
			xhr = new XMLHttpRequest();
			xhr.open("HEAD", url, false);
			xhr.send(null);
			if(xhr.status == 404) {
				url = slide_dir + '/slide00.jpg';
			}
			$.toast({
				 text : '<div class="h7">' + url + '</div>'
				,loader: false
				,bgColor : 'white'			// 背景色
				,textColor : 'blue'			// 文字色
				,position : 'bottom-right'	// ページ内での表示位置
				,hideAfter : 10 * 1000
			});
		    $('.centerWindow').css({
				backgroundImage: 'url("' + url + '")'
			});
			$('.centerWindow').fadeIn('slow',function(){
//				chime();
				$(this).delay(slide_dly * 1000).fadeOut('slow');
			});
	    },slide_int * 1000);
	}
	$('.centerWindow').on('click', function(){
		$(this).hide();
	});
    $('#select_config').change(function() {
//		alert($(this).val());
		$.toast({text: this.id + ':' + $(this).val(), loader: false});
		switch($(this).val()) {
		case 'w7':
			setConfig('#JGYOBU'			, 'J'		);
			setConfig('#UKEHARAI_CODE'	, 'ZI0'		);
			setConfig('#sumi_disp'		, 'false'	);
			setConfig('#p_sshiji_chg'	, '10'		);
			setConfig('#p_shorder_chg'	, '1'		);
			setConfig('#p_sagyo_log_chg', '1'		);
			setConfig('#zaiko9_chg'		, '1'		);
			$('.config').trigger('change');
			break;
		}
    });
});
/*
fetchWindow
*/
$(document).ready(function() {
	fetchWait('');
});
function fetchWait(id) {
	if( id == '') {
		$('#fetchWindow').text(id)
		return false;
	}
	id = '#' + id;
	if( $('#fetchWindow').text() != '') {
		$('#fetchWindow').text($('#fetchWindow').text() + '.');
		var wait = ($('#fetchWindow').text().match(/\./g)||[]).length;
		console.log(id + '...wait...' + $('#fetchWindow').text() + wait);
/*
		$.toast({
			 text : '<div class="h6">' + id + '...wait...' + $('#fetchWindow').text() + wait + '</div>'
			,bgColor : 'black'
			,textColor : 'white'
			,loader: false
			,hideAfter : 10000
		});
*/
		setTimeout(function() {
			$(id).trigger("click");
		},wait * 1000);
		return true;
	}
	$('#fetchWindow').text(id)
	return false;
}
function utimeOffset(utime, cont) {
	console.log('utimeOffset(' + utime + ', ' + cont + ')');
//	console.log('.offset().top: ' + $(utime).offset().top);
//	console.log('.offset().left: ' + $(utime).offset().left);
	var	left = $(cont + ' > table').offset().left + $(cont + ' > table').width();
	if(left > $(window).width()) {
		left = $(window).width();
	}
	left -= $(utime).width();
	$(utime).offset({
		 top: $(window).scrollTop() + $(cont).offset().top - $(utime).height()
		,left: left
	});
}
/*
出荷状況 y_syuka_h
*/
$(document).ready(function() {
	setConfig('#y_syuka_h_chg','');
	setConfig('#y_syuka_h_scr','');
	setConfig('#y_syuka_h_fetch',600);
//	setConfig('#y_syuka_h_to',0);
	$('#y_syuka_h_div').on('focus', function() {
		$('#navChg').text($('#y_syuka_h_chg').val());
		$('#navScr').text($('#y_syuka_h_scr').val());
	});
	var timer = null;
	$('#y_syuka_h_update').on('click', function(){
//		if($('#navId').text() == '#y_syuka_h_div' || $(this).text() == '◇') {
		// 同時処理防止
		if(fetchWait(this.id) == false) {
			$(this).text('■');
			//出荷状況更新
			debug_toast('<div class="h7">click:' + this.id + '</div>');
			$('a[href~="#y_syuka_h_div"]').text('出荷状況');
			var	url = 'monitor.py?f=y_syuka_h&dns=' + $('#dns').val();
			if($('#KAN_DT').val() != '') {
				url += ('&KAN_DT=' + $('#KAN_DT').val());
			}
			//検索実行
			debug_toast('<div class="h7">fetch:' + url + '</div>');
			fetch(url)
			.then( function ( res ) {
				// 同時処理解除
				fetchWait('');
				$('#y_syuka_h_update').text('□');
				var contentType = res.headers.get("content-type");
				debug_toast('<div class="h7">' + url + ':' + contentType + '</div>');
				if(contentType && contentType.indexOf("application/json") !== -1) {
					console.log('return json');
					return res.json();
				} else {
					$('#y_syuka_h_update').text('Error');
					$.toast({
						text : 'y_syuka_h error:' + contentType + ' ' + url
						,loader: false
						,hideAfter : 60000
					});
				}
			} )
			.then( function ( json ) {
				var now = new Date();
	//			$('#tm').html(datetimeFormat(now) + '<span class="text-yellow">更新</span>');
	//			debug_toast('<div class="h7">json.data.length:' + json.data.length + '</div>');
	//			debug_toast('<div class="h7">json.sql:' + json.sql + '</div>');
				var	tr = '';
				var tDlvQty = 0;
				var tCnt = 0;
				var tPicZan = 0;
				var tCntZan = 0;
				var	tTr = null;
				for ( var i = 0 ; i < json.data.length ; i++ ) {
		//			debug_toast('<div class="h7">json.data[' + i + ']</div>');
					switch(json.data[i].UNSOU_KAISHA) {
					case '福山通運':
					case '日本通運':
					case '久留米運送':
					case '第一貨物':
					case '西武航空':
					case 'ヤマト':
					case 'PANA':
						tDlvQty += json.data[i].DlvQty;
						tCnt += json.data[i].Cnt;
						tPicZan += json.data[i].PicZan;
						tCntZan += json.data[i].CntZan;
						break;
					default:
						if(!tTr && tDlvQty > 0) {
							tTr = '';
							tr += '<tr>';
							tr += '<td class="h1 text-right">運送会社計</td>';
							tr += '<td></td>';
							tr += cnt_td("h0",tDlvQty);
							tr += cnt_td("h0",tCnt);
							tr += cnt_td("h2 zan",tPicZan);
							tr += cnt_td("h2 zan",tCntZan);
							tr += '</tr>';
							tDlvQty = 0;
						}
						break;
					}
					tr += '<tr>';
					tr += '<td class="h0">' + json.data[i].UNSOU_KAISHA + '</td>';
					//便
					tr += '<td>';
					var	j = 0;
					[json.data[i].DlvQty1
					,json.data[i].DlvQty2
					,json.data[i].DlvQty3
					,json.data[i].DlvQty9].forEach(function( q ) {
						j++;
						if(q > 0) {
							tr += '<div>';
							tr += '<span class="bin">' + (j == 4 ? 9 : j) + '</span>';
							tr += q;
							tr += '</div>';
						}
					});
					tr += '</td>';
					//送り状枚数
					tr += '<td>' + json.data[i].DlvQty + '</td>';
					tr += cnt_td("h0",json.data[i].Cnt);
	//				tr += '<td class="h0 text-right">' + json.data[i].DlvQty + '</td>';
	//				tr += '<td class="h0 text-right">' + json.data[i].Cnt + '</td>';
					switch(json.data[i].UNSOU_KAISHA.trim()) {
					case '積水':
					case 'ミサワ':
						tr += '<td class="h2 text-right"></td>';
						tr += '<td class="h2 text-right"></td>';
						break;
					default:
						tr += cnt_td("h2 zan",json.data[i].PicZan);
						tr += cnt_td("h2 zan",json.data[i].CntZan);
	//					tr += '<td class="h2 text-right">' + json.data[i].PicZan + '</td>';
	//					tr += '<td class="h2 text-right">' + json.data[i].CntZan + '</td>';
						break;
					}
					tr += '</tr>';
				}
				$('#y_syuka_h > tbody').find("tr").remove();
				$('#y_syuka_h > tbody').append(tr);
				$('#y_syuka_h_cancel').change();
				$('#y_syuka_h_update').text(nowTM());
			} );
		}
		$('#y_syuka_h_cancel').change(function(){
			debug_toast(this.id + '.change()');
			var	url = 'monitor.py?f=y_syuka_h_cancel&dns=' + $('#dns').val();
			if($('#KAN_DT').val() != '') {
				url += ('&KAN_DT=' + $('#KAN_DT').val());
			}
			//未出庫リスト
			url += '&UnPick=1';
			//検索実行
			fetch(url)
			.then( function ( res ) {
				var contentType = res.headers.get("content-type");
				debug_toast('<div class="h7">' + url + ':' + contentType + '</div>');
				if(contentType && contentType.indexOf("application/json") !== -1) {
					return res.json();
				} else {
					$.toast({
						text : this.id + ' error:' + contentType + ' ' + url
						,loader: false
						,hideAfter : 60000
					});
				}
			} )
			.then( function ( json ) {
				debug_toast(url + ':' + json.data.length);
				$('#y_syuka_h_cancel > tbody').find("tr").remove();
				if(json.data.length == 0) {
					return;
				}
				var	tr = '';
				var	cancel = json.data.reduce((a,x) => a += Number(x.CANCEL_F),0);
				if(cancel > 0) {
					tr += '<tr class="bg-white"><td colspan="4">キャンセル：' + cancel + '件</td></tr>';
				}
				for ( var i = 0 ; i < json.data.length ; i++ ) {
					if(i < cancel) {
						tr += '<tr class="bg-white">';
						tr += '<td>' + json.data[i].INS_BIN + '</td>';
						tr += '<td>' + json.data[i].MUKE_NAME + '</td>';
						tr += '<td>' + json.data[i].HIN_NO + '</td>';
						tr += '<td>' + json.data[i].Qty + '</td>';
						tr += '</tr>';
					} else {
						if(i == cancel) {
							tr += '<tr class="bg-gray"><td colspan="4">未出庫：' + (json.data.length - cancel) + '件';
							if((json.data.length - cancel) > 9) {
								tr += '　※10件未満でリストを表示します.</td></tr>';
								break;
							}
							tr += '</td></tr>';
						}
						tr += '<tr class="bg-gray">';
						tr += '<td>' + json.data[i].INS_BIN + '</td>';
						tr += '<td>' + json.data[i].MUKE_NAME + '</td>';
						tr += '<td>' + json.data[i].HIN_NO + '</td>';
						tr += '<td>' + json.data[i].Qty + '</td>';
						tr += '</tr>';
					}
				}
				if(zaiko9_data.length > 0) {
					var	qty90 = 0;
					for ( var i = 0 ; i < zaiko9_data.length ; i++ ) {
						if(zaiko9_data[i].Soko == '90') {
							qty90++;
						}
					}
					tr += '<tr class="bg-black"><td colspan="4">仮置き残(90)：' + qty90 + '件</td></tr>';
					tr += '<tr class="bg-black"><td colspan="4">.</td></tr>';
					tr += '<tr class="bg-black"><td colspan="4">.</td></tr>';
					tr += '<tr class="bg-black"><td colspan="4">.</td></tr>';
				}

				$('#y_syuka_h_cancel > tbody').append(tr);
			} );
		});
		if($('#y_syuka_h_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					debug_toast('setTimeout:#y_syuka_h:' + $('#navId').text());
					$('#y_syuka_h_update').trigger("click");
				},$('#y_syuka_h_fetch').val() * 1000);
			}
		}
		return false;
	});
	if($('#y_syuka_h_chg').val() > 0) {
		$('#y_syuka_h_update').trigger("click");
		$('#y_syuka_h_div').removeClass('disable');
	} else {
		$('#y_syuka_h_div').addClass('disable');
	}
});
/*
出荷状況 y_syuka
*/
$(document).ready(function() {
	setConfig('#y_syuka_chg','');
	setConfig('#y_syuka_scr','');
	setConfig('#y_syuka_fetch',600);
	setConfig('#y_syuka_to','');
	$('#y_syuka_div').on('focus', function() {
		$('#navChg').text($('#y_syuka_chg').val());
		$('#navScr').text($('#y_syuka_scr').val());
//		$(this).children("table").floatThead('destroy');
		utimeOffset('#y_syuka_update','#' + this.id);
//		$(this).children("table").floatThead({top: $(this).offset().top});
	});
//	console.log('#y_syuka > thead > tr:eq(1) > th:eq(1).css("display"):'
//				 + $("#y_syuka > thead > tr:eq(1) > th:eq(1)").css("display"));
//	if( $("#y_syuka > thead > tr:eq(1) > th:eq(1)").css("display") == "none") {
//		$("#y_syuka > thead > tr:eq(0) > th:eq(1)").attr("colSpan","1");
//	}
	var timer = null;
	$('#y_syuka_update').on('click', function(){
//		if($('#navId').text() == '#y_syuka_div' || $(this).text() == '◇') {
		$(this).text($(this).text().replace('□','■'));
		// 同時処理防止
		if(fetchWait(this.id) == false) {
			//出荷状況更新
			debug_toast('<div class="h7">click:' + this.id + '</div>');
			$('a[href~="#y_syuka_div"]').text('出荷状況');
			var	url = 'monitor.py?f=y_syuka&dns=' + $('#dns').val();
			if($('#JGYOBU').val() != '') {
				url += ('&JGYOBU=' + $('#JGYOBU').val());
			}
			if($('#KAN_DT').val() != '') {
				url += ('&KAN_DT=' + $('#KAN_DT').val());
			}
			//検索実行
			debug_toast('<div class="h7">fetch:' + url + '</div>');
			fetch(url)
			.then((res) => {
				// 同時処理解除
				fetchWait('');
				var contentType = res.headers.get("content-type");
				debug_toast('<div class="h7">' + url + ':' + contentType + '</div>');
				if(contentType && contentType.indexOf("application/json") !== -1) {
					console.log('return json');
					return res.json();
				} else {
					$.toast({
						text : 'y_syuka error:' + contentType + ' ' + url
						,loader: false
						,hideAfter : 60000
					});
				}
			} )
			.then((json) => {
				var now = new Date();
	//			$('#tm').html(datetimeFormat(now) + '<span class="text-yellow">更新</span>');
				var	tr = '';
				var	curr = null;
				var	prev = null;
				var	stat2 = false;
				for ( var i = 0 ; i < json.data.length ; i++ ) {
					curr = td_date(json.data[i].KEY_SYUKA_YMD);
//					if(json.data[i].Stts == '3') {
//						curr += ' (引当済)';
//						curr += '<span class="h4"> (伝発待ち)</span>';
//					}
					if(curr != prev) {
						var	c = "future";
						var	zan = '';
						if(prev == null) {
							c = "today";
							var	Zan = 0;
							var	Zan9 = 0;
							var	ZanA = 0;
							var	ZanA9 = 0;
							for ( var j = i ; j < json.data.length ; j++ ) {
								if(json.data[i].KEY_SYUKA_YMD != json.data[j].KEY_SYUKA_YMD) {
									break;
								}
								if(json.data[j].Stts != '4') {
									break;
								}
								Zan += json.data[j].CntZan9;
								Zan9 += json.data[j].CntZan;
								ZanA += json.data[j].CntZanLK;
							}
							zan += '<div class="box t"> 出庫残</div><div class="box' + (Zan > 0 ? ' zan' : '') + '">' + Zan + '</div>';
							zan += '<div class="box t"> 検品残</div><div class="box' + (Zan9 > 0 ? ' zan' : '') + '">' + Zan9 + '</div>';
							zan += '<div class="box t"> Active送信残</div><div class="box' + (ZanA > 0 ? ' zan' : '') + '">' + ZanA + '</div>';
//							zan += '<span class="h4"> 実績残：' + ZanA9 + ')</span>';
						}
						tr += '<tr><td class="' + c + '">' + curr + '</td>';
						tr += '<td class="' + c + '" colspan="10">' + zan + '</td></tr>';
					}
					prev = curr;
					switch(json.data[i].DestCode) {
					case 'A3':
					case 'A6':
					case 'A7':
						stat2 = true;
					}
					if(json.data[i].Stts == '3') {
						tr += '<tr class="stts3">';
					} else {
						tr += '<tr class="stts4 ' + json.data[i].DestCode + '" title="' + json.data[i].DestCode + '">';
					}
					tr += '<td class="dest"><div class="dest">' + json.data[i].Dest + '</div></td>';
//					tr += '<td class="dest">' + json.data[i].Dest + '</td>';
					//合計 件数
					tr += cnt_td("Cnt",json.data[i].Cnt);
					//合計 個数
					tr += cnt_td("total-qty",json.data[i].Qty);
					//合計 才数
					tr += cnt_td("total-sai",(json.data[i].Sai).toFixed(1));
					//合計 出庫残
//					tr += cnt_td("h1 zan",json.data[i].CntZan9);
//					if(json.data[i].CntS > 0) {
//						tr = tr.slice(0,-5);
//						tr += '<br><span class="h7">未商品:</span>' + json.data[i].CntS + '</td>';
//					}
					//合計 検品残
//					tr += cnt_td("h1 zan",json.data[i].CntZan);
					//合計 送信残
//					tr += cnt_td("h3",json.data[i].CntZanLK);
					//補充
					if(json.data[i].Cnt3 > 0) {
						tr += cnt_td("Cnt",json.data[i].Cnt3);
						if(json.data[i].Stts == '3') {
							tr += cnt_td("h3",json.data[i].Cnt3Zan9);	
							tr += cnt_td("h3",json.data[i].Cnt3Zan);
						} else {
							tr += cnt_td("h3 zan",json.data[i].Cnt3Zan9);	
							tr += cnt_td("h3 zan",json.data[i].Cnt3Zan);
						}
					} else {
						tr += '<td></td>';
						tr += '<td></td>';
						tr += '<td></td>';
					}
					//緊急
					if(json.data[i].Cnt2 > 0) {
						tr += cnt_td("Cnt",json.data[i].Cnt2);
						if(json.data[i].Stts == '3') {
							tr += cnt_td("h3",json.data[i].Cnt2Zan9);
							tr += cnt_td("h3",json.data[i].Cnt2Zan);
						} else {
							tr += cnt_td("h3 zan",json.data[i].Cnt2Zan9);
							tr += cnt_td("h3 zan",json.data[i].Cnt2Zan);
						}
					} else {
						tr += '<td></td>';
						tr += '<td></td>';
						tr += '<td></td>';
					}
					tr += '</tr>';
				}
				if($(window).scrollTop() == 0) {
//2018.12.28		$('#y_syuka_div > table').floatThead('destroy');
				}
				$('#y_syuka > tbody').find("tr").remove();
				$('#y_syuka > tbody').append(tr);
				$(this).text('□更新.' + nowTM());
				utimeOffset('#' + this.id, '#y_syuka_div');
//				$('#y_syuka_div > table').floatThead({top: $('#y_syuka_div').offset().top});
	//			var	top = $('#y_syuka').offset().top;
	//			$('#y_syuka').floatThead({
	//							top: top,
	//							position: 'auto',	//'auto','fixed','absolute'
	//							});
	//			$('#y_syuka').floatThead('reflow');
				if(stat2 == true) {
//					$('#stat2').removeClass('disable');
//					$('#stat2_update').trigger("click");
//					$('#y_spot').removeClass('disable');
//					$('#y_spot_update').trigger("click");
				}
			} )
//			.catch(function(err) {
			.catch((err) => {
				console.error(err);
				$.toast({text: '<div class="h7">' + url + '</div><div class="h7">fetch err:' + err + '</div>',loader: false, hideAfter : 60000});
			});
		}
		if($('#y_syuka_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					debug_toast('setTimeout:#y_syuka:' + $('#navId').text());
					$('#y_syuka_update').trigger("click");
				},$('#y_syuka_fetch').val() * 1000);
			}
		}
		return false;
	});
	if($('#y_syuka_chg').val() > 0) {
		$('#y_syuka_update').trigger("click");
		$('#y_syuka_div').removeClass('disable');
	} else {
		$('#y_syuka_div').addClass('disable');
	}
});

function cnt_td(cls,num) {
	cls += ' number';
	if(cls.indexOf('zan') != -1) {
		if(num > 0) {
			cls += ' bg-zan';
		}
	}
	return '<td class="' + cls + '">' + num + '</td>';
}
/*
海外出荷状況
*/
$(document).ready(function() {
	setConfig('#order_chg','');
	setConfig('#order_scr','');
	setConfig('#order_fetch',600);
	setConfig('#order_chg_pm','');
	setConfig('#order_scr_pm','');
	setConfig('#order_fetch_pm','');
	setConfig('#order_to','');
	$('#order').on('focus', function() {
		$('#navChg').text($('#order_chg').val());
		$('#navScr').text($('#order_scr').val());
		var now = new Date();
		var	tm = now.getHours() * 100 + now.getMinutes();
		if (tm > 1200) {
			if($('#order_chg_pm').val() != '') {
				$('#navChg').text($('#order_chg_pm').val());
			}
			if($('#order_scr_pm').val() != '') {
				$('#navScr').text($('#order_scr_pm').val());
			}
		}
		utimeOffset('#order_update','#' + this.id);
//		$(this).children('table').floatThead({top: $(this).offset().top});
	});
	var timer = null;
	$('#order_update').on('click', function(){
//		if($('#navId').text() == '#order' || $(this).text() == '◇') {
		$(this).text($(this).text().replace('□','■'));
		// 同時処理防止
		if(fetchWait(this.id) == false) {
			var	url = 'monitor.py?f=order&dns=' + $('#dns').val();
			console.log('fecth():' + url);
			debug_toast('<div class="h7">' + url + '</div>');
			//検索実行
			fetch(url)
			.then((res) => {
				// 同時処理解除
				fetchWait('');
				var contentType = res.headers.get("content-type");
				debug_toast('<div class="h7">' + url + ':' + contentType + '</div>');
				if(contentType && contentType.indexOf("application/json") !== -1) {
					console.log('return json');
					return res.json();
				} else {
					$.toast({
						text : 'order error:' + contentType + ' ' + url
						,loader: false
						,hideAfter : 60000
					});
				}
			} )
			.then((json) => {
				debug_toast('<div class="h7">order:json:' + json.data.length + '</div>');
				// 海外出荷予定
				var	tr = '';
				var today = dateYMD(new Date());
				for ( var i = 0 ; i < json.data.length ; i++ ) {
					debug_toast(i + ':' + json.data[i].MUKE_NAME);
					var	cls = '';
					if(json.data[i].KEY_SYUKA_YMD != json.data[i].KEY_SYUKA_YMD_0) {
						cls += "senko ";
					}
					if(today == json.data[i].KEY_SYUKA_YMD) {
						cls += "today ";
					} else {
					}
					tr += '<tr class="' + cls + '">';

					tr += '<td class="h4 text-right">' + (i + 1) + '</td>';
					var dt = json.data[i].KEY_SYUKA_YMD.slice(-4);
					dt = Number(dt.slice(0,2)) + '/' + Number(dt.slice(-2));
					tr += '<td class="text-center" title="' + json.data[i].KEY_SYUKA_YMD_0 + '">' + dt + '</td>';
					tr += '<td class="h4 text-left">' + json.data[i].MUKE_NAME + '</td>';
					tr += '<td class="h4 text-left">' + json.data[i].CYU_KBN_NAME + '</td>';
					tr += '<td title="order_no" class="h4 text-left order_no">' + json.data[i].ODER_NO + '</td>';
					tr += '<td class="text-right">' + json.data[i].cnt + '</td>';
					if(json.data[i].KEY_SYUKA_YMD != json.data[i].KEY_SYUKA_YMD_0) {
						var dt = json.data[i].KEY_SYUKA_YMD_0.slice(-4);
						dt = Number(dt.slice(0,2)) + '/' + Number(dt.slice(-2));
						tr += '<td class="h5 text-center" colspan="2">先行出荷(' + dt + ')</td>';
					} else {
						if(json.data[i].zan0 == 0) {
							tr += '<td class="text-center text-morning">済</td>';
						} else {
							tr += '<td class="text-right">' + json.data[i].zan0 + '</td>';
						}
						if(json.data[i].zan9 == 0) {
							tr += '<td class="text-center text-morning">済</td>';
						} else {
							tr += '<td class="text-right">' + json.data[i].zan9 + '</td>';
						}
					}
					tr += '<td class="text-right">' + json.data[i].qty + '</td>';
					tr += '</tr>';
				}
				console.log('fecth:end');
				if($(window).scrollTop() == 0) {
//2018.12.28		$('#order_list').floatThead('destroy');
				}
				$('#order_list > tbody').find("tr").remove();
				$('#order_list > tbody').append(tr);
				$('#order_sum').trigger("click");
				$(this).text('□更新.' + nowTM());
				utimeOffset('#' + this.id, '#order');
				
			});
		}
		if($('#order_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					debug_toast('setTimeout:#order:' + $('#navId').text());
					$('#order_update').trigger("click");
				},$('#order_fetch').val() * 1000);
			}
		}
		return false;
	});
	var order_chg = $('#order_chg').val();
	var now = new Date();
	var	tm = now.getHours() * 100 + now.getMinutes();
	if (tm > 1200) {
		if($('#order_chg_pm').val() != '') {
			order_chg = $('#order_chg_pm').val();
			if(order_chg == 0) {
				$.toast({
					text : 'PM～海外出荷状況は非表示に設定されています.'
					,loader: false
					,hideAfter : 20000
				});
			}
		}
	}

	if(order_chg > 0) {
		$('#order_update').trigger("click");
		$('#order').removeClass('disable');
	} else {
		$('#order').addClass('disable');
	}
	// 出荷日別の件数／残数
	$('#order_sum').on('click', function(){
		var thead = '<tr><th>出荷日</th>';
		var tbody1 = '<tr><th>件数</th>';
		var tbody2 = '<tr><th>残</th>';
		var ym = '';
		var ym_prev = '';
		var	daycnt = 0;
		var	zan = 0;
		var tr = $("#order_list > tbody > tr");
		var	cnt = 0;
		var	cls = '';
		for(var i = 0; i < tr.length; i++ ){
			if(ym != $('td',tr.eq(i)).eq(1).text()) {
				if(ym != '') {
					thead += '<td class="' + cls + '">' + ym + '</td>';
					tbody1 += '<td class="' + cls + '">' + cnt + '</td>';
					tbody2 += '<td class="' + cls + '">' + zan + '</td>';
					cls = '';
				}
				ym = $('td',tr.eq(i)).eq(1).text();
				cnt = 0;
				zan = 0;
			}
			if(i > 0) {
				if($('td',tr.eq(i)).eq(1).text() != $('td',tr.eq(i-1)).eq(1).text()) {
					daycnt = 0;
				}
				
				if($('td',tr.eq(i)).eq(1).text() == $('td',tr.eq(i-1)).eq(1).text() &&
				   $('td',tr.eq(i)).eq(2).text() == $('td',tr.eq(i-1)).eq(2).text() ) {
					$('td',tr.eq(i)).eq(1).addClass('hidden');
					$('td',tr.eq(i)).eq(2).addClass('hidden');
				} else {
				}
			}
			if((i + 1) >= tr.length) {
				$('td',tr.eq(i)).addClass('daybottom');
			} else if($('td',tr.eq(i)).eq(1).text() != $('td',tr.eq(i+1)).eq(1).text()) {
				$('td',tr.eq(i)).addClass('daybottom');
			}
			daycnt++;
			$('td',tr.eq(i)).eq(0).text(daycnt);
			if(daycnt > 1) {
				if($('td',tr.eq(i)).eq(2).text() == $('td',tr.eq(i-1)).eq(2).text()) {
					$('td',tr.eq(i)).eq(2).addClass('hidden');
					if($('td',tr.eq(i)).eq(3).text() == $('td',tr.eq(i-1)).eq(3).text()) {
						$('td',tr.eq(i)).eq(3).addClass('hidden');
					}
				}
			}
			console.log($('td',tr.eq(i)).eq(1).text() + ';' + $('td',tr.eq(i)).hasClass('today'));
			if($(tr.eq(i)).hasClass('today')) {
				cls = 'today';
			}
			cnt += GetNumber($('td',tr.eq(i)).eq(5).text());
			if($('td',tr.eq(i)).eq(6).text().search( /^先行出荷.*$/ ) == 0) {
				//先行出荷
			} else {
				zan += GetNumber($('td',tr.eq(i)).eq(7).text());
			}
		}
		if(ym != '') {
			thead += '<td class="' + cls + '">' + ym + '</td>';
			tbody1 += '<td class="' + cls + '">' + cnt + '</td>';
			tbody2 += '<td class="' + cls + '">' + zan + '</td>';
		}
		thead += '</tr>';
		tbody1 += '</tr>';
		tbody2 += '</tr>';
//		$('#order_sum > thead').find("tr").remove();
//		$('#order_sum > thead').append(thead);
		$('#order_sum').find("tr").remove();
		$('#order_sum').append(thead + tbody1 + tbody2);
//		$('#order_list').floatThead({top: $('#order').offset().top});
		return false;
	});


	//出荷日変更
//	$('#order_list > tbody > tr > td').on('click', function() {
//	$(document).on("click","#order_list > tbody > tr",function() { 
//	$(document).on('click', '.order_no', function() {
//	$('tr').on('click', '#order_list > tbody', function() {
//	$(document).on('click', '#order_list', function() {
//	$('.order_no').on('click', function() {
//	$('td').on('click', '.order_no', function() {
//	$(document).on('click', '#order_list td', function() {
//	$(document).on('click', '#order_list tbody', function() {
//	$(document).on('click', '#order_list > tbody', function() {
//	$('#order_list > tbody > tr').on('click', function() {
//	$(document).on('click', '#order_list td', function() {
//	$(document).on('click', '#order_list', function() {
//	$('#order_list > tbody').on('click', function() {	//ok
//	$('#order_list').on('click', function() {	//ok
//	$('#order_list td').on('click', function() {
//	$('#order_list tbody').on('click', function() {	//ok
//	$('#order_list tbody td').on('click', function() {
//	$('#order_list').on('click', 'td', function() {	//ok
	$('#order_list').on('click', '.order_no', function() {	//ok
//		console.log('click:' + this.id);
		$tag_td = $(this)[0];
		$tag_tr = $(this).parent()[0];
//		console.log("%s列, %s行", $tag_td.cellIndex, $tag_tr.rowIndex);
//		alert('click:' + $(this).text() + ' ' + $('td',$(this).parent()).eq(1).attr('title'));
		$("#eKEY_SYUKA_YMD").val($('td',$(this).parent()).eq(1).attr('title'));
		$("#eMUKE_NAME").val($('td',$(this).parent()).eq(2).text());
		$("#eCYU_KBN_NAME").val($('td',$(this).parent()).eq(3).text());
		$("#eODER_NO").val($(this).text());
		$("#ecnt").val($('td',$(this).parent()).eq(5).text());
		$("#eqty").val($('td',$(this).parent()).eq(8).text());
        $("#dialog-edit").dialog("open");
        return false;
	});
    $("#dialog-edit").dialog({
        autoOpen: false,
        height: 'auto',
        width: 'auto',
        modal: true,
        buttons: {  // ダイアログに表示するボタンと処理
			"変更": function() {
				//出荷日変更
				var	url = 'syukadt.py?dns=' + $('#dns').val();
				url += '&ODER_NO=' + $('#eODER_NO').val();
				url += '&SYUKA_YMD=' + $('#eKEY_SYUKA_YMD').val();
				console.log(url);
				fetch(url)
				.then((res) => {
					var contentType = res.headers.get("content-type");
					console.log(contentType);
					$.toast({
						 text : '<div class="h7">' + url + ': ' + contentType + '</div>'
						,bgColor : 'white'
						,textColor : 'black'
						,loader: false
						,hideAfter : 5000
					});
					if(contentType && contentType.indexOf("application/json") !== -1) {
						return res.json();
					}
					if(contentType && contentType.indexOf("text/html") !== -1) {
						return res.text();
					}
				})
				.then((json) => {
					console.log(json);
					$(this).dialog("close");
				})
				.then((text) => {
					console.log(text);
					$.toast({
						 text : '<div class="h7">' + '出荷日を変更しました.' + $('#eKEY_SYUKA_YMD').val() + '</div>'
						,bgColor : 'white'
						,textColor : 'black'
						,loader: false
						,hideAfter : 5000
					});
					$('#order_update').trigger("click");
					$(this).dialog("close");
				});
			},
            "キャンセル": function() {
				$(this).dialog("close");
            }
        },
        // ダイアログのイベント処理
        open: function(event, ui) {
//			console.log('dialog.open():' + this.id);
			// タイトル 編集：行番号 品番
			$(this).dialog('option', 'title', '出荷日変更：' + $('#eODER_NO').val());
        },
        close: function() {
        }
    });
});

function GetNumber(x){ 
	if(isNumber(x)) {
		return Number(x);
	}
	return 0;
}
function isNumber(x){ 
    if( typeof(x) != 'number' && typeof(x) != 'string' )
        return false;
    else 
        return (x == parseFloat(x) && isFinite(x));
}
/*
商品化状況 p_sshiji
*/
var	p_sshiji_data = [];
$(document).ready(function() {
	setConfig('#p_sshiji_chg','');
	setConfig('#p_sshiji_scr','');
	setConfig('#p_sshiji_fetch',600);
	$('#package').on('focus', function() {
//		$('#navCnt').text(0);
		$('#navChg').text($('#p_sshiji_chg').val());
		$('#navScr').text($('#p_sshiji_scr').val());
		utimeOffset('#p_sshiji_update','#' + this.id);
//		return false;
	});
//	setConfig('#p_sshiji_to',0);
	setConfig('#UKEHARAI_CODE','');
	setConfig('#KAN_DT','');
	if(/AC|AC./.test($('#UKEHARAI_CODE').val())) {
		$('th',$('#summary > tbody > tr').eq(0)).eq(1).text('ＳＸ');
		$('th',$('#summary > tbody > tr').eq(0)).eq(2).text('出荷分');
		$('th',$('#summary > tbody > tr').eq(0)).eq(3).text('当日出荷分');
	}
	var timer = null;
	$('#p_sshiji_update').on('click', function(){
//		if($('#navId').text() == '#package' || $(this).text() == '◇') {
		$(this).text($(this).text().replace('□','■'));
		// 同時処理防止
		if(fetchWait(this.id) == false) {
			var	f = 'p_sshiji';
//			var	url = 'monitor.py?dns=' + $('#dns').val();
			var	url = 'p_sshiji.py?dns=' + $('#dns').val();
			debug_toast('title:' + $('a[href~="#package"]').text());
			$('a[href~="#package"]').text('商品化状況');
			if($('#UKEHARAI_CODE').val() != '') {
				$('a[href~="#package"]').text('商品化状況：' + $('#UKEHARAI_CODE').val());
				url += ('&UKEHARAI_CODE=' + $('#UKEHARAI_CODE').val());
				if($('#UKEHARAI_CODE').val() == 'ZD3') {
					f = 'p_sshiji_92';
				}
			}
			if($('#JGYOBU').val() != '') {
				url += ('&JGYOBU=' + $('#JGYOBU').val());
			}
			if($('#KAN_DT').val() != '') {
				url += ('&KAN_DT=' + $('#KAN_DT').val());
			}
			url += '&f=' + f
			//検索実行
			debug_toast('<div class="h7">' + url + '</div>');
			fetch(url)
			.then((res) => {
				// 同時処理解除
				fetchWait('');
				var contentType = res.headers.get("content-type");
				debug_toast('<div class="h7">' + url + ': ' + contentType + '</div>');
				if(contentType && contentType.indexOf("application/json") !== -1) {
					console.log('return json');
					return res.json();
				} else {
					$.toast({
						text : 'p_sshiji error:' + contentType + ' ' + url
						,loader: false
						,hideAfter : 60000
					});
				}
			} )
			.then((json) => {
	//			var now = new Date();
	//			$('#tm').html(datetimeFormat(now) + '<span class="text-yellow">更新</span>');
				p_sshiji_data = json.data;
				if(p_sshiji_data.length == 0) {
					$.toast({text: '商品化状況：該当データがありません.'
							,loader: false});
					return;
				}
				$('#summary > tbody > tr > td').find('span').text(0);
				$('#summary > tbody > tr > td').find('div').text(0);
				var	tr = '';
				for ( var i = 0 ; i < json.data.length ; i++ ) {
					var	status = '';
					var	shiji_f = json.data[i].SHIJI_F;
					var	bg = 'bg-black';
					//0:通常 1:ｽﾎﾟｯﾄ 2：欠品解除 3:再梱包
					if(/AC|AC./.test($('#UKEHARAI_CODE').val())) {
						switch( shiji_f ) {
						case '0':	shiji_f = 'ＳＸ';		bg = 'jizen';	break;
						case '1':	shiji_f = '出荷分';		bg = 'spot';	break;
						case '2':	shiji_f = '当日出荷分';	bg = 'keppin';	break;
						case '3':	shiji_f = '再梱包';			break;
						}
					} else {
						switch( shiji_f ) {
						case '0':	shiji_f = '事前';		bg = 'jizen';	break;
						case '1':	shiji_f = 'スポット';	bg = 'spot';	break;
						case '2':	shiji_f = '欠品解除';	bg = 'keppin';	break;
						case '3':	shiji_f = '再梱包';		break;
						}
					}
					var	flash = '';
					switch(kanCheck(json.data[i])) {
					case 9:	//9 完了
						tr += '<tr class="h4 bg-black">';
						status = '済' + '<div class="h5">' + json.data[i].KAN_DT + '</div>';
						break;
					case 8:	//8 完了
						tr += '<tr class="h4 bg-black">';
						status = '済' + '<div class="h5">92出庫(' + json.data[i].sumiQty92 + ')</div>';
						break;
					case 1:	//1 商品化チェック済
						tr += '<tr class="h4 ' + bg + '">';
						status = 'チェック';
						status += '<div class="h5">' + json.data[i].HIN_CHECK_DATETIME.slice(0,8) + '</div>';
						break;
					case 0:	//0 予定
						status = '';
						tr += '<tr class="h4 ' + bg + '">';
						break;
					}
//					if ( status.indexOf('済') < 0 ) {
					if ( status == '') {
						if(json.data[i].sumiQty92 != 0){
							status += '<div class="h5">92出庫<br>(' + json.data[i].sumiQty92 + ')</div>';
						}
					}
					tr += '<td class="text-right">' + (i + 1) + '</td>';
//					var dt = json.data[i].HAKKO_DT.slice(-4);
					var	dt = '';
					if(typeof json.data[i].YOTEI_DT !== 'undefined') {
						dt = json.data[i].YOTEI_DT.slice(-4);
						if(dt != '') {
							dt = Number(dt.slice(0,2)) + '/' + Number(dt.slice(-2));
						}
					}
					tr += '<td class="text-center">' + dt + '</td>';
					//納期回答日(AcNoki)
					dt = '';
					if(typeof json.data[i].AcNoki !== 'undefined') {
						dt = json.data[i].AcNoki;	//.slice(-4);
//						if(dt != '') {
//							dt = Number(dt.slice(0,2)) + '/' + Number(dt.slice(-2));
//						}
					}
					tr += '<td class="text-center">' + dt + '</td>';
					tr += '<td class="text-center">'
					tr += shiji_f;
					if (json.data[i].yQty > 0) {
						tr += '<div class="h5">';
						var	yQty = json.data[i].yQty - json.data[i].yQty2;
						if(json.data[i].yQty2 > 0) {
							tr += '<div class="box stts2">欠品</div><div class="box stts2">' + json.data[i].yQty2 + '</div>';
						}
						if(yQty > 0) {
							if(yQty >= json.data[i].zqtySumi) {
								tr += '\n<div class="box">出荷</div><div class="box">' + yQty + '</div>';
							} else {
								tr += '\n<div class="box stts0">出荷</div><div class="box stts0">' + yQty + '</div>';
							}
						}
						tr += '</div>';
					}
					tr += '</td>'
					tr += '<td class="text-left">' + json.data[i].SHIJI_NO;
					tr += '<div class="h6">' + json.data[i].HAKKO_DT + ' ' + json.data[i].UKEHARAI_CODE + '</div>';
					tr += '</td>';
					tr += '<td class="text-left h0 hover">' + json.data[i].HIN_GAI;
					tr += '<div class="tooltip">';
					tr += '<p>' + json.data[i].HIN_GAI + '</p>'
					tr += '<p>' + json.data[i].HIN_NAME + '</p>'
					tr += '<p>' + json.data[i].SType + '</p>'
					tr += '<p>' + json.data[i].SCat + '</p>'
					tr += '</div>';
					tr += '</td>';
					tr += '<td class="text-right h0">' + json.data[i].qty + '</td>';
					tr += '<td class="text-center">';
					if(json.data[i].HIN_CHECK_DATETIME != '') {
						tr += '<div class="check icon"></div>';
					}
					tr += status;
					tr += '</td>';
					//在庫92
					tr += '<td>';
					if(json.data[i].zqty92 > 0) {
						tr += json.data[i].zqty92;
					} else {
						if(typeof json.data[i].zqty92today !== 'undefined') {
							if(json.data[i].zqty92today > 0) {
								tr += '(' + json.data[i].zqty92today + ')';
							}
						}
					}
					tr += '</td>';
					//在庫未
					tr += '<td>';
					if(json.data[i].zqtyMi > 0) {
						tr += json.data[i].zqtyMi;
					}
					tr += '</td>';
					//在庫済
					tr += '<td>';
					if(json.data[i].zqtySumi > 0) {
						tr += json.data[i].zqtySumi;
					}
					tr += '</td>';
					tr += '</tr>';
					countSummary(json.data[i]);
				}
			//	tr += '<tr class="h4 text-dummy"></tr>';
//				if($(window).scrollTop() == 0) {
//					$('#package > table').floatThead('destroy');
//				}
				$('#tbody_p_sshiji').find("tr").remove();
				$('#tbody_p_sshiji').append(tr);
				tbody_p_sshiji_sumi();
				$(this).text('□更新.' + nowTM());
				utimeOffset('#' + this.id, '#package');
				console.log(this.id + '.text():' + $(this).text());
//				$('#package > table').floatThead({top: $('#package > table').offset().top});
/*
				$('#p_sshiji').floatThead({
									top: $('#package').offset().top,
									position: 'auto',	//'auto','fixed','absolute'
									});
*/
				console.log('json.message:' + json.message);
				if(typeof json.message !== 'undefined') {
					if(json.message != '') {
						$.toast({text: '商品化状況<p>' + json.message
								,bgColor : 'white'
								,textColor : 'blue'
								,loader: false
								,hideAfter : 20 * 1000
								});
					}
				}
				//商品化状況(納期別) JCS
//				var	noki = $('#p_sshiji > tbody > tr:first > td:eq(2)').text().trim();
				console.log('json.data[0].AcNoki:' + json.data[0].AcNoki);
				var	noki = '';
				if(typeof json.data[0].AcNoki !== 'undefined') {
					noki = '1';
				}
				console.log(this.id + '.noki:' + noki);
				if(noki != '') {
//					$.toast('商品化状況(納期別):ON:' + noki);
					$('#p_sshiji_jcs').removeClass('disable');
					$('#p_sshiji_jcs_update').trigger("click");
				} else {
//					$.toast('商品化状況(納期別):OFF:' + noki);
					$('#p_sshiji_jcs').addClass('disable');
				}
//				if(typeof json.data[0].SCat !== 'undefined') {
//					$('#p_sshiji_cat').removeClass('disable');
//					$('#p_sshiji_cat_update').trigger("click");
//				}
			} );
		}
		if($('#p_sshiji_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					debug_toast('setTimeout:#package:' + $('#navId').text());
					$('#p_sshiji_update').trigger("click");
				},$('#p_sshiji_fetch').val() * 1000);
			}
		}
		return false;
	});
	if($('#p_sshiji_chg').val() > 0) {
		$('#p_sshiji_update').trigger("click");
		$('#package').removeClass('disable');
	} else {
		$('#package').addClass('disable');
	}
	//tooltip
//	$('td').hover(function() {
//	$('td').on({
	$(document).on('mouseenter', '.hover', function(){
		console.log('td.mouseenter ' + $(this).text());
		$(this).children('.tooltip').show();
	});
	$(document).on('mouseleave', '.hover', function(){
		console.log('td.mouseleave ' + $(this).text());
		$(this).children('.tooltip').hide();
	});
});
// 商品化済を非表示
function tbody_p_sshiji_sumi() {
	var	tr = $('#p_sshiji > tbody > tr');
	num = 0;
	for( i = 0 ; i < tr.length ; i++) {
		console.log('tbody_p_sshiji_sumi():' + i + ':' + $('td',tr.eq(i)).eq(6).text());
		$('td',tr.eq(i)).eq(0).text(++num);
		if( $('td',tr.eq(i)).eq(7).text().indexOf('済') >= 0 ) {
//			$.toast({text: 'tbody_p_sshiji_sumi():' + i + ':' + $('td',tr.eq(i)).eq(6).text(),loader: false,hideAfter : 5000});
			console.log('tbody_p_sshiji_sumi():remove:' + $('#sumi_disp').text());
			if($('#sumi_disp').val() == 'false') {
				$(tr.eq(i)).remove();
				num--;
			}
		}
	}
/*
	if($('#sumi_disp').val() == 'false') {
		$.toast({text: '商品化済を非表示にしました.',loader: false,hideAfter : 10000});
	}
*/
}
//
function countSummary(p_sshiji) {
	var	f = p_sshiji.SHIJI_F;
	//0:通常 1:ｽﾎﾟｯﾄ　2：欠品解除 3:再梱包
	switch( f ) {
	case '0':	countUp($('#f0_y'),1);	break;	//'通常';
	case '1':	countUp($('#f1_y'),1);	break;	//'ｽﾎﾟｯﾄ';
	case '2':	countUp($('#f2_y'),1);	break;	//'欠品解除';
	case '3':	countUp($('#f3_y'),1);	break;	//'再梱包';
	default:	return;	break;
	}
	countUp($('#total_y'),1);
	switch(kanCheck(p_sshiji)) {
	case 9:	//9 完了
	case 8:	//8 完了
		countUp($('#f' + f + '_k'),1);
		countUp($('#total_k'),1);
		break;
	case 1:	//1 商品化中
		countUp($('#f' + f + '_s'),1);
		countUp($('#total_s'),1);
		break;
	case 0:	//0 予定
	}
	$('#f' + f + '_z').text(Number($('#f' + f + '_y').text()) - Number($('#f' + f + '_k').text()));
	$('#total_z').text(Number($('#total_y').text()) - Number($('#total_k').text()));
	countUp($('#f' + f + '_q'),Number(p_sshiji.qty));
	countUp($('#total_q'),Number(p_sshiji.qty));
}
function kanCheck(p_sshiji) {
	if( p_sshiji.KAN_F == '1' ) {
		//9 完了
		return 9;
	}
	if(p_sshiji.zqty92 == 0) {
		if (p_sshiji.sumiQty92 != 0) {
			//8 完了 92から出庫済
			return 8;
		}
	}
	//1 商品化チェック済
	if(p_sshiji.HIN_CHECK_TANTO) {
		return 1;
	} else {
		//商品化未チェック
/*
		if(p_sshiji.zqty92 == 0) {
			if(p_sshiji.GOODS_YMD >= p_sshiji.HAKKO_DT) {
				//おまけ
				//9 完了
				return 8;
			}
		}
*/
	}
	//0 予定
	return 0;
}
function countUp(c,num) {
	num += Number($(c).text());
	$(c).text(num);
}
// dateFormat 関数の定義
function datetimeFormat(date) {
  var y = date.getFullYear();
  var m = date.getMonth() + 1;
  var d = date.getDate();
  var w = date.getDay();
  var wNames = ['日', '月', '火', '水', '木', '金', '土'];
  var hh = date.getHours();
  var mm = date.getMinutes();
/*
  if (m < 10) {
    m = '0' + m;
  }
  if (d < 10) {
    d = '0' + d;
  }
*/
  if (mm < 10) {
    mm = '0' + mm;
  }

  // フォーマット整形済みの文字列を戻り値にする
  return y + '.' + m + '.' + d + '(' + wNames[w] + ') ' + hh + ':' + mm;
}
function nowTM() {
	var now = new Date();
	return datetimeFormat(now).slice(-5);
}
function dateYMD(date, c = '') {
  var y = date.getFullYear();
  var m = date.getMonth() + 1;
  var d = date.getDate();
  if (m < 10) {
    m = '0' + m;
  }
  if (d < 10) {
    d = '0' + d;
  }
  // フォーマット整形済みの文字列を戻り値にする
  return '' + y + c + m + c + d;
}
/*
Chime音

chrome://flags/#autoplay-policy
No user gesture is required

発車メロディＭＩＤＩ
http://www47.tok2.com/home/cs381/hassya.html

midi→mp3
https://www.conversion-tool.com/midi?lang=en
*/
function chime(nm = 'chime', volume = 0.1) {
	console.log('chime()' + nm, + ',' + volume);
/*
	$.toast({
		 text : 'chime():' + nm
		,loader: false
		,hideAfter : 3 * 1000
	});
*/
	audio = new Audio();
	audio.src = 'sound/' + nm + '.mp3';
	audio.load();
	if($('#volume').val()) {
		audio.volume = parseInt($('#volume').val());
	} else {
		audio.volume = volume;
	}
//	if(audio.paused) {
		audio.play().catch(function(e) {
			console.log('audio.play():' + e);
			$.toast({
				 heading: 'audio.play().error'
				,text : '<div class="h5">' + audio.src + '</div>' + '<div class="h5">' + e + '</div>' + '<input type="button" onClick="audio.play();" value="♪">'
				,loader: false
				,hideAfter : 30 * 1000
			});
		});
//	}
/*
	sound = new Audio('sound/' + nm + '.mp3');
	sound.load();
	sound.play();
*/
/*
	$('#' + nm)[0].play();
*/
/*
	var $audio = $(nm).get(0);
	$audio.volume = 0.2;
	$audio.play();
*/
/*
	var audio = document.getElementById(nm);
	audio.play();
*/
}
/*
注意メッセージ表示
*/
$(document).ready(function() {
//	chime();
	$.toast({
//		heading : ''
		text : 'Pos作業モニター <span class="h4">' + version + '</span>'	//表示したいテキスト(HTML使用可)
		,showHideTransition : 'slide'	// 表示・消去時の演出
		,bgColor : 'white'				// 背景色
		,textColor : 'black'			// 文字色
		,allowToastClose : true			// 閉じるボタンの表示・非表示
		,hideAfter : 10 * 1000			// 自動的に消去されるまでの時間(ミリ秒)(falseを指定すると自動消去されない)
//		,stack : 5						// 一度に表示できる数
		,textAlign : 'left'				// テキストの配置
		,position : 'bottom-left'		// ページ内での表示位置
		,loader: false
//		,loaderBg: 'cyan'
//		,icon: 'info'
	});

	$('#morning').on('click', function(){
		text = eq();
		$.toast({
			text : text
//			,heading : '4/5'
			,bgColor : 'white'
			,textColor : 'darkblue'
			,textAlign : 'center'
			,position : 'bottom-left'
			,loader: false
			,hideAfter : 60 * 1000
		});
//		console.log("$('#morning_play').trigger('click')");
	});
	$('#slogan2018').on('click', function(){
		$.toast({
			text : '<div>2018年度スローガン</div>' +
				   '<div class="h0">強い意志で目標を達成する</div>' +
				   '<div>社長</div>'
			,textAlign : 'center'
			,bgColor : 'white'
			,textColor : 'black'
			,position : 'bottom-left'
			,loader: false
			,hideAfter : 30 * 1000
		});
	});
	$('#slogan2019').on('click', function(){
		$.toast({
			text : '<div>2019年度スローガン</div>' +
				   '<div class="h0">謙虚に驕らずチャレンジする</div>' +
				   '<div>社長</div>'
			,textAlign : 'center'
			,bgColor : 'white'
			,textColor : 'black'
			,position : 'bottom-left'
			,loader: false
			,hideAfter : 30 * 1000
		});
	});
	$('#slogan').on('click', function(){
		$.toast({
			text : '<div>2020年度スローガン</div>' +
				   '<div class="h0">衆知を集め 全員経営</div>' +
				   '<div>社長</div>'
			,textAlign : 'center'
			,bgColor : 'white'
			,textColor : 'black'
			,position : 'bottom-left'
			,loader: false
			,hideAfter : 30 * 1000
		});
	});
	$('#evening').on('click', function(){
		$.toast({
			text : '<div class="h0">終業時間です。お疲れ様でした。' +
				   '気を付けて帰宅して下さい。' +
				   '最終退出者は火の元・戸締りを確認して下さい。</div>'
			,bgColor : 'white'
			,textColor : 'black'
			,position : 'bottom-left'
			,loader: false
			,hideAfter : 60 * 1000
		});
	});
	$('#endofyear').on('click', function(){
		$.toast({
			text : '<div class="endofyear"><div>本年も大変お世話になり、ありがとうございました。 m(_ _)m<p>' +
				   'よいお年をお迎えください。^o^/<p></div><div class="h4">所長は火の元・戸締りの最終確認を報告願います。</div></div>'
			,bgColor : 'pink'
			,textColor : 'black'
			,position : 'bottom-left'
			,loader: false
			,hideAfter : 180 * 1000
		});
	});
	$('#care1').on('click', function(){
//		chime('care1');
		chime('se_maoudamashii_jingle03');
		$.toast({
			heading : '<div class="h0 text-center text-red blinking">≪誤出荷注意≫</div>'
//<!--caption class="">誤出荷発生状況 ※全社</caption>
			,text : '<center><table class="gosyuka">'
					+ '<tr><th colspan="3">誤出荷発生状況</th><th>小野</th><th>滋Ｐ</th><th>滋物</th><th>袋井</th><th>大阪</th><th>奈良</th><th>広島</th></tr>'
					+ '<tr><td>2020年度</td><td class="text-red">１件</td><td class="text-left h4">(6/18現在)</td>'
					+ '<td>－</td>'
					+ '<td>－</td>'
					+ '<td>－</td>'
					+ '<td>－</td>'
					+ '<td>－</td>'
					+ '<td class="text-red">１</td>'
					+ '<td>－</td>'
					+ '</tr>'
					+ '<tr><td>2019年度</td><td>１５件</td><td class="text-left h4">(年間累計)</td><td>４</td><td>２</td><td>２</td><td>－</td><td>１</td><td>４</td><td>２</td></tr>'
					+ '<tr><td>2018年度</td><td>１０件</td><td class="text-left h4">(年間累計)</td><td>－</td><td>３</td><td>１</td><td>１</td><td>１</td><td>４</td><td>－</td></tr>'
					+ '</table></center>'
					+ '<div class="box">誤出荷は１拠点だけの問題でなく、会社全体の信用・信頼を損ないます。'
					+ '細心の注意を心がけましょう！</div>'
					+ '<div class="h2 text-red -blinking marquee">'
					+ '2020年度発生分'
					+ '  ①6/17 奈良センター倉庫：送り先違い'
					+ '</div>'
//					+ '2019年度発生分'
//					+ '  ⑮11/18 奈良センター倉庫：中身違い(セット品不足 )'
//					+ '　⑭10/23 広島事：ラベル違い'
//					+ '　⑬10/19 奈良センター倉庫：送り先違い'
//					+ '　⑫10/15 大阪事業所：送り先違い'
//					+ '　⑪10/3 滋賀DC：中身違い(ラベル貼り間違い)'
//					+ '　⑩9/24 滋賀PC：中身違い'
//					+ '　⑨7/31 小野PC：その他(個装ラベル未貼付)'
//					+ '　⑧6/28 広島事：数量違い'
//					+ '　⑦6/24 小野PC：数量違い'
//					+ '　⑥6/15 奈良センター倉庫：送り先違い'
//					+ '　⑤6/14 奈良センター倉庫：その他（現品違い）'
//					+ '　④6/13 滋賀PC：送り先違い'
//					+ '　③6/7 小野PC：数量違い'
//					+ '　②6/7 滋賀PC：数量違い'
//					+ '　①5/21 小野PC：納品書誤送'
//					+ '</div>'
			,bgColor : 'white'
			,textColor : 'black'
			,position : 'bottom-left'
			,loader: false
			,hideAfter : 60 * 1000
		});
	});
	$('#gensan1').on('click', function(){
		$.toast({
			text : '<div class="h0"><span class="strong">原産国</span>は<span class="strong">パーツラベル</span>・<span class="strong">アイテムラベル</span>・<span class="strong">送り状<span class="h2">(Packing List)</span></span>の３点チェック！'
			,heading : '<div class="h0 text-center text-red">海外出荷梱包時の注意事項</div>'
			,bgColor : 'lemonchiffon'		//'orange'
			,textColor : 'black'
			,position : 'bottom-left'
			,loader: false
			,hideAfter : 60 * 1000
		});
	});
	$('#gensan2').on('click', function(){
		var	text = '<div class="h0 text-center">｢目視可能な範囲」で現物確認</div>';
		text += '<div class="h0 text-center"><span class="box text-red">原産国<div class="h7">　　　　　　</div></span>';
		text += '<span> </span>';
		text += '<span class="box text-blue">品番<div class="h7">(現物刻印有のみ)</div></span>';
		text += '<span> </span>';
		text += '<span class="box text-green">色<div class="h7">　　　　　　</div></span>';
		text += '<span> </span>';
		text += '<span class="box">形状<div class="h7">　　　　　　</div></span>';
		text += '</div>';
		$.toast({
			 heading : '<div class="h0 text-center text-red">商品化時の注意事項</div>'
//			,text : '<div class="h0"><span class="strong">原産国</span>は<span class="strong">現物表示</span>と<span class="strong">パーツラベル</span>を確認！</div>'
			,text : text
			,bgColor : 'lemonchiffon'
			,textColor : 'black'
			,position : 'bottom-left'
			,loader: false
			,hideAfter : 60 * 1000
		});
	});
	$('#noon').on('click', function(){
		$.toast({
			text : '<div class="h0">春の全国交通安全運動</div>' +
				   '<div>4/6～4/15 実施中です。</div>' +
				   '<div>交通ルールを守り安全運転に努めましょう。</div>'
			,textAlign : 'center'
			,bgColor : 'lightgreen'
			,textColor : 'black'
			,position : 'bottom-left'
			,loader: false
			,hideAfter : 30 * 1000
		});
	});
	$('#safe').on('click', function(){
		$.toast({
			text : '<div>12月</div>' +
				   '<div class="h0">『安全強化月間』</div>' +
				   '<table class="text-left text-blue">' +
				   '<tr><td>　　　　</td><td>・人の安全行動基準の唱和</td></tr>' +
				   '<tr><td></td><td>・フォークリフト安全運転</td></tr>' +
				   '<tr><td></td><td>・火の元消火の確実な確認、トラッキング予防</td></tr>' +
				   '<tr><td></td><td>・通勤/営業中の車両安全運転…早めのライト点灯</td></tr>' +
				   '</table>' +
				   '<p>道交法の改訂で「ながら運転」の罰則が強化されました。</p>'
			,textAlign : 'center'
			,bgColor : 'lightyellow'
			,textColor : 'black'
			,position : 'bottom-left'
			,loader: false
			,hideAfter : 90 * 1000
		});
	});
	$('#order_care').on('click', function(){
		// 同時処理防止
		if(fetchWait(this.id)) {
			return;
		}
		//※海外出荷《注意事項》
		var	url = 'order.py?dns=' + $('#dns').val();
		if($('#KAN_DT').val() != '') {
			url += '&table=' + $('#KAN_DT').val();
		}
		fetch(url)
		.then((res) => {
			// 同時処理解除
			fetchWait('');
			return res.json();
		})
		.then((json) => {
			if(json.data.length == 0) {
				return;
			}
			var text = '';
			text += '<table class="toast">'
			for ( var i = 0 ; i < json.data.length ; i++ ) {
				text += '<tr>'
				text += '<td>' + json.data[i].KEY_SYUKA_YMD + '</td>';
				text += '<td>' + json.data[i].MUKE_NAME;
				text += '<div class="h4">' + json.data[i].LK_MUKE_CODE + '</div></td>';
				text += '<td>' + json.data[i].ODER_NO + '</td>';
				text += '<td>' + json.data[i].KEY_HIN_NO;
				text += '<div class="HIN_NAME">' + json.data[i].iHIN_NAME + '</div></td>';
				text += '<td class="text-right">' + json.data[i].qty + '</td>';
				text += '<td class="careful">' + json.data[i].Careful + '</td>';
				text += '</tr>'
			}
			text += '</table>'
			$.toast({
//							heading : json.data.length
				 text : '※海外出荷《注意事項》' + text
				,bgColor : 'yellow'
				,textColor : 'black'
				,loader: false
				,hideAfter : 15000
			});
		});
	});

	setInterval(function(){
		toastMessage();
    },400 * 1000);			//間隔 200000ms 200秒 3分20秒
	//test用 すぐに実行
	toastMessage();
});
function eq() {
//			text : '<div class="h0">新年あけましておめでとうございます。今年もよろしくお願い申し上げます。</div>' 	//+ now.getMonth()
//			text : '<div class="h0">おはようございます。今日も１日元気にがんばりましょう!</div>'
//			text : '<div class="h0">2018年度(52期)新年度のスタートです.</div>'
	var	array = [];
	array.push('<div class="h0">今日は一日、笑顔でいこう.</div><div class="h3">まず鏡の前で笑顔筋をマッサージ、笑顔はあなたの感情を前向きにします.</div>');
	array.push('<div class="h0">今日は意識して<br>積極的に振る舞おう.</div><div class="h3">「私がやります！」－前向きな感情は積極的な行動をつくります.</div>');
	array.push('<div class="h0">今日は10人を誉めよう.</div><div class="h3">誉める気持ちをつくりだすことで、あなたのこころも強くなります.</div>');
	array.push('<div class="h0">今日は明るい言葉を使って話そう.</div>'
			 + '<div class="h3">「明るい言葉」はあなたも、あなたの周囲も元気にさせます.</div>');
	array.push('<div class="h0">今日の「得した気分」を味わおう。</div>'
			 + '<div class="h3">「なんか得したなあ」は感動の一歩です。<p>「得したなあ」の数ほど、こころの感度が磨かれます。</div>');
	array.push('<div class="h0">今日はスキップをしてみよう。</div>'
			 + '<div class="h3">最近、スキップしていますか。楽しい気分は楽しく振る舞うことで生まれます。</div>');
	array.push('<div class="h0">今日のランチメニューは何ですか。</div>'
			 + '<div class="h3">誰かに合わせるのでなく、たまには自分の感情にわがままになってみましょう。</div>');
	array.push('<div class="h0">今日は相手のこころが<p>躍ることをしよう。</div>'
			 + '<div class="h3">相手が思いっきり喜ぶことをしてみましょう。あなたも楽しい気分になれます。</div>');
	array.push('<div class="h0">今日は思いっきり<p>楽しいことを考えよう。</div>'
			 + '<div class="h3">楽しい気分は、あなたのこころの栄養素です。今日の夢はきっと叶います。</div>');
	array.push('<div class="h0">今日は自分の感情の<p>動きをメモしてみよう。</div>'
			 + '<div class="h3">感情に注目することができると感情に賢く対応できるようになります。</div>');
	array.push('<div class="h0">今日は人の話を最後まで聞く日です。</div>'
			 + '<div class="h3">とにかく我慢して聞きましょう。こころが鍛えられます。</div>');
	array.push('<div class="h0">今日は10人を励まそう。</div>'
			 + '<div class="h3">相手を励ますことで、あなたのこころも元気になります。</div>');
	array.push('<div class="h0">今日の「感動した！」は何ですか。</div>'
			 + '<div class="h3">一日一回感動しましょう。<p>感動の数だけ人を感動させることができます。<p>感動できるあなたのこころの感度は絶好調です。</div>');
	array.push('<div class="h0">今日は自分自身を誉めてあげよう。</div>'
			 + '<div class="h3">一番嬉しくなる言葉で自分を誉めてあげましょう。<p>あなたにエネルギーのプレゼントです。</div>');
	array.push('<div class="h0">今日は拍手の日。</div>'
			 + '<div class="h3">いいことに出会ったら思いっきり拍手をしましょう。<p>あなたにも周囲にもエネルギーが生まれます。</div>');
	array.push('<div class="h0">今日は10人と握手をしよう。</div>'
			 + '<div class="h3">握手は心の距離を近くします。<p>こころの距離感を体感してください。</div>');
	array.push('<div class="h0">今日はあなたの大切な人の<p>気分を聞いてみよう。</div>'
			 + '<div class="h3">人の気分や気持ちがわかれば、言葉や行動は的確になります。</div>');
	array.push('<div class="h0">今日の朝一番は、<p>楽しい会話で始めよう。</div>'
			 + '<div class="h3">楽しい会話は、今日一日の前向きな感情をつくります。</div>');
	array.push('<div class="h0">今日はクヨクヨの日。</div>'
			 + '<div class="h3">クヨクヨ気分は成長の源です。失敗は解決できる人のところにやってくる！</div>');
	array.push('<div class="h0">今日は３秒決断の日。</div>'
			 + '<div class="h3">迷わず即決の日です。「決める！」快感をたのしみにしましょう。</div>');
	array.push('<div class="h0">「今日はツイてる！」<p>と言葉にして言おう。</div>'
			 + '<div class="h3">「ツイてる！」の言葉であなたのこころにエネルギーを注入しましょう。</div>');
	array.push('<div class="h0">今日は10人にあいさつをしよう。</div>'
			 + '<div class="h3">あいさつは相手の感情をオープンにします。人間関係はあいさつから始まります。</div>');
	array.push('<div class="h0">今日は「感謝しています」を<p>10人に言おう。</div>'
			 + '<div class="h3">感謝とは相手のこころに敬意を払う気持ちです。言葉にして伝えましょう。</div>');
	array.push('<div class="h0">「あー、つまんないなあ」<p>今の気分を意識的に感じてみよう。</div>'
			 + '<div class="h3">気分は変えることができます。「やる気」は自由自在につくれるのです。</div>');
	array.push('<div class="h0">今日は大笑いの日。</div>'
			 + '<div class="h3">人に笑われるくらいの笑顔で笑ってみましょう。こころも体も軽くなります。</div>');
	array.push('<div class="h0">何もしない<p>空白の時間をつくろう。</div>'
			 + '<div class="h3">時間や仕事に縛られない自由時間は、疲れたこころを癒してくれます。</div>');
	array.push('<div class="h0">やる気が出ないときは、<p>5分だけ背筋を伸ばしてみよう。</div>'
			 + '<div class="h3">背筋を伸ばし、胸を張って歩きましょう。気分が変わります。</div>');
	array.push('<div class="h0" title="29">今日は自分のこころの<p>踊ることをしよう。</div>'
			 + '<div class="h3">エネルギー充電の日です。思いっきり奮発して美味しいものを食べましょう。</div>');
	array.push('<div class="h0" title="30">今日は一人の時間を楽しもう。</div>'
			 + '<div class="h3">今夜は自分の感情との対話を楽しみ、<p>自分のこころと素敵なデートをしてみましょう。</div>');
	array.push('<div class="h0" title="31">今日は怒らない日にしよう。</div>'
			 + '<div class="h3">ムカッときたら「シックスセカンズ」。６秒間待てば、こころが静まります。</div>');
	var day1 = new Date("2018/05/22");
	var day2 = new Date();
//	var termDay = Math.ceil((day2 - day1) / 86400000);
	var termDay = Math.floor((day2 - day1) / (1000 * 60 * 60 *24));
	return array[termDay % array.length];
}
function toastMessage() {
	var now = new Date();
	console.log('toastMessage():' + now + ' getHours:' + now.getHours());
	var	dt = (now.getMonth() + 1) * 100 + now.getDate();
	var	tm = now.getHours() * 100 + now.getMinutes();
	//朝の挨拶
	if (tm <= 930) {
		$('#morning').trigger('click');
		return;
	}
	if (tm > 1200 && tm < 1245) {
		if(dt >= 406 && dt <= 415) {
			$('#noon').trigger('click');
		}
	}
	//終業の挨拶
	if (tm > 1700) {
		if(dt >= 1227) {
			$('#endofyear').trigger('click');
			return;
		}
		$('#evening').trigger('click');
		return;
	}
	//誤出荷注意
//	if(now.getMinutes() == 20) {
//		$('#care1').trigger('click');
//		return;
//	}
	//スローガン
	if(now.getMinutes() == 15) {
		$('#slogan').trigger('click');
		return;
	}
	//月間
	if(now.getMinutes() > 57) {
		if(dt >= 1201 && dt <= 1231) {
			$('#safe').trigger('click');
			return;
		}
	}
	//出荷状況(大阪)が有効な場合は原産国注意を非表示
	if($('#y_syuka_h_chg').val() > 0) {
//		$('#zaiko90').trigger('click');
		return ;
	}
	if($('#order_chg').val() > 0) {
		if((now.getMinutes() % 2) == 0) {
			if($('#order_chg_pm').val() == '') {
				$('#order_care').trigger('click');
				return;
			} else {
				var	tm = now.getHours() * 100 + now.getMinutes();
				if(tm < 1200) {
					$('#order_care').trigger('click');
					return;
				}
			}
		}
	}
	//原産国注意
	switch($('#navId').text()) {
	case '#order':
		$('#gensan1').trigger('click');
		break;
	case '#package':
		$('#gensan2').trigger('click');
		break;
	}
	return ;
}
//nav
$.fn.navChange = function(first) {
	console.log('navChange():' + first);
//	debug_toast('navChange():' + $('a',this).length);
	if(!first && $('#title').text() == '設定') {
		return false;
	}
	var t = '';
	var	curr = null;
	var	next = null;
//	for(i = 0; i < $(this).children().length; i++) {
//		child = $(this).children().eq(i);
	$('a',this).each(function(i) {
		console.log('navChange():' + $(this).text() + ':' + $(this).attr('href') + ':' + $($(this).attr('href')).attr("class"));
//		debug_toast(i + '/' + $('a','nav').length + ':' + $(this).text() + ':' + $(this).attr('href'));
		var	child = this;
		var	target = $(child).attr('href');
		if(!($(target).hasClass('disable'))) {
			if(first) {
				$(child).trigger("click");
				console.log('navChange():first:' + first);
				return false;
			}
			if(!next) {
				next = child;
				title = $(child).text();
			}
			if(curr) {
				next = child;
				title = $(child).text();
				$(next).trigger("click");
				debug_toast('return:' + i + ':' + $(next).text() + ':' + $(next).attr('href'));
				return false;
			}
		}
		debug_toast('navChange():' + target + ':' + $(target).is(':visible'));
		if($(target).is(':visible')) {
			if(!first && target == '#config_div') {
				return false;
			}
			curr = child;
		}
		if(i == $('a','nav').length - 1 ) {
			//最後
			$(next).trigger("click");
			debug_toast('last:' + i + ':' + $(next).text() + ':' + $(next).attr('href'));
			return false;
		}
	});
};
/*
資材入荷予定 p_shorder
*/
$(document).ready(function() {
	setConfig('#p_shorder_chg','');
	setConfig('#p_shorder_scr','');
	setConfig('#p_shorder_fetch',600);
	$('#p_shorder').on('focus', function() {
		$('#navChg').text($('#p_shorder_chg').val());
		$('#navScr').text($('#p_shorder_scr').val());
		utimeOffset('#p_shorder_update','#' + this.id);
	});
	var timer = null;
	$('#p_shorder_update').on('click', function(){
//		if($('#navId').text() == '#p_shorder' || $(this).text() == '◇') {
		$(this).text($(this).text().replace('□','■'));
		// 同時処理防止
		if(fetchWait(this.id) == false) {
			var	url = 'monitor.py?f=p_shorder&dns=' + $('#dns').val();
			if($('#JGYOBU').val() != '') {
				url += ('&JGYOBU=' + $('#JGYOBU').val());
			}
			//検索実行
			debug_toast('<div class="h7">' + url + '</div>');
			fetch(url)
			.then((res) => {
				// 同時処理解除
				fetchWait('');
				var contentType = res.headers.get("content-type");
				debug_toast('<div class="h7">p_shorder:' + contentType + '</div>');
				if(contentType && contentType.indexOf("application/json") !== -1) {
					return res.json();
				} else {
					$.toast({
						text : 'p_shorder error:' + contentType + ' ' + url
						,loader: false
						,hideAfter : 60000
					});
				}
			} )
			.then((json) => {
				debug_toast('<div class="h7">p_shorder:json:' + json.data.length + '</div>');
				var	tr = '';
				for ( var i = 0 ; i < json.data.length ; i++ ) {
					tr += '<tr>';
					tr += '<td class="h5 number">' + (i + 1) + '</td>';
					tr += '<td class="date">' + td_date(json.data[i].KAN_DT) + '</td>';
					tr += '<td class="date">' + td_date(json.data[i].Y_NOUKI_DT) + '</td>';
					tr += '<td class="">' + json.data[i].HIN_GAI + '</td>';
					tr += '<td class="number">' + json.data[i].qty + '</td>';
					tr += '<td class="number">' + json.data[i].zQty + '</td>';
					tr += '<td class="h5">' + json.data[i].ORDER_CODE + '</td>';
					tr += '<td class="h5">' + json.data[i].G_SYUSHI + '</td>';
					tr += '</tr>';
				}
				if($(window).scrollTop() == 0) {
//2018.12.28		$('#p_shorder > table').floatThead('destroy');
				}
				$('#p_shorder > table > tbody').find("tr").remove();
				$('#p_shorder > table > tbody').append(tr);
				$(this).text('□更新.' + nowTM());
				utimeOffset('#' + this.id, '#p_shorder');
//				$('#p_shorder > table').floatThead({top: $('#p_shorder').offset().top});
			} );
		}
		if($('#p_shorder_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					debug_toast('setTimeout:#p_shorder:' + $('#navId').text());
					$('#p_shorder_update').trigger("click");
				},$('#p_shorder_fetch').val() * 1000);
			}
		}
		return false;
	});
	$('a[href~="#p_shorder"]').text('資材入荷予定');
	if($('#p_shorder_chg').val() > 0) {
		$('#p_shorder_update').trigger("click");
	} else {
		$('#p_shorder').addClass('disable');
	}
});
function td_date(dt) {
	if(dt == '') {
		return dt;
	}
	var dt4 = dt.slice(-4);
	return Number(dt4.slice(0,2)) + '/' + Number(dt4.slice(-2));
}
/*
資材補充アラーム s_alarm
*/
$(document).ready(function() {
	setConfig('#s_alarm_chg','');
	setConfig('#s_alarm_scr','');
	setConfig('#s_alarm_fetch',600);
	$('#s_alarm').on('focus', function() {
		$('#navChg').text($('#s_alarm_chg').val());
		$('#navScr').text($('#s_alarm_scr').val());
	});
	var timer = null;
	utimeOffset('#s_alarm_update','#s_alarm');
	$('#s_alarm_update').on('click', function() {
		$(this).text('◆' + $(this).text().slice(1));
		if(fetchWait(this.id) == false) {
			$(this).text('■' + $(this).text().slice(1));
			var	url = 's_alarm.py?dns=' + $('#dns').val();
			fetch(url)
			.then((res) => {
				// 同時処理解除
				fetchWait('');
				var contentType = res.headers.get("content-type");
				debug_toast('<div class="h7">s_alarm:' + contentType + '</div>');
				if(contentType && contentType.indexOf("application/json") !== -1) {
					return res.json();
				} else {
					$.toast({
						text : 's_alarm error:' + contentType + ' ' + url
						,loader: false
						,hideAfter : 60000
					});
					return;
				}
			} )
			.then((json) => {
				debug_toast('<div class="h7">s_alarm:json:' + json.data.length + '</div>');
				var	tr = '';
				for ( var i = 0 ; i < json.data.length ; i++ ) {
					tr += '<tr>';
					tr += '<td class="rowNo">' + (i + 1) + '</td>';
					tr += '<td class="HIN_GAI">' + json.data[i].HIN_GAI;
					tr += '<div class="HIN_NAME">' + json.data[i].HIN_NAME + '</div>';
					tr += '</td>';
					tr += '<td class="Qty">' + getValue(json.data[i].zqty) + '</td>';
					tr += '<td class="Qty">' + getValue(json.data[i].HOJYU_P) + '</td>';
					tr += '<td class="Qty">' + getValue(json.data[i].zanQty) + '</td>';
					tr += '<td class="Qty';
					if (parseInt(json.data[i].Fit2Qty) < 0) {
						tr += ' Short';
					}
					tr += '">';
					tr += getValue(json.data[i].Fit2Qty) + '</td>';
					tr += '<td class="Date">' + getValue(json.data[i].maxKAN_DT) + '</td>';
					tr += '<td class="G_SYUSHI">' + json.data[i].G_SYUSHI;
					tr += '<div class="C_NAME">' + json.data[i].C_NAME + '</div>';
					tr += '</td>';
					tr += '</tr>';
				}
				tr += '<tr>';
				tr += '<td class="rowNo"></td>';
				tr += '<td class="" colspan="' + ($('#s_alarm > table > thead').find("th").length - 1) + '">';
				tr += (json.data.length) + '件';
				tr += '</td>';
				tr += '</tr>';
				$('#s_alarm > table > tbody').find("tr").remove();
				$('#s_alarm > table > tbody').append(tr);
				$(this).text('□更新.' + nowTM());
				utimeOffset('#s_alarm_update','#s_alarm');
//				$('#p_shorder > table').floatThead({top: $('#p_shorder').offset().top});
			} );
		}
		if($('#s_alarm_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					$('#s_alarm_update').trigger("click");
				},$('#s_alarm_fetch').val() * 1000);
			}
		}
		return false;
	});
	if($('#s_alarm_chg').val() > 0) {
		$('#s_alarm_update').trigger("click");
		$('#s_alarm').removeClass('disable');
	} else {
		$('#s_alarm').addClass('disable');
	}
});
function getValue(v) {
	return (v ? v : '');
}
/*
作業状況 p_sagyo_log
*/
$(document).ready(function() {
	setConfig('#p_sagyo_log_chg',10);
	setConfig('#p_sagyo_log_scr','');
	setConfig('#p_sagyo_log_fetch',600);
	$('#p_sagyo_log').on('focus', function() {
//		$('#navCnt').text(0);
		$('#navChg').text($('#p_sagyo_log_chg').val());
		$('#navScr').text($('#p_sagyo_log_scr').val());
		$('#p_sagyo_log_update').offset({
			 top: $('#p_sagyo_log > table').offset().top - $('#p_sagyo_log_update').height()
			,left: $('#p_sagyo_log > table').offset().left + $('#p_sagyo_log > table').width() - $('#p_sagyo_log_update').width()
		});
//		return false;
	});
	var timer = null;
	$('#p_sagyo_log_update').on('click', function(){
//		if($('#navId').text() == '#p_sagyo_log' || $(this).text() == '◇') {
		$(this).text($(this).text().replace('□','■'));
		// 同時処理防止
		if(fetchWait(this.id) == false) {
//			$(this).text('■');
			debug_toast('click:' + this.id);
			var	url = 'monitor.py?f=p_sagyo_log&dns=' + $('#dns').val();
			if($('#JGYOBU').val() != '') {
				url += ('&JGYOBU=' + $('#JGYOBU').val());
			}
			if($('#KAN_DT').val() != '') {
				url += ('&JITU_DT=' + $('#KAN_DT').val());
			}
			//検索実行
			debug_toast('<div class="h7">' + url + '</div>');
			fetch(url)
			.then((res) => {
				// 同時処理解除
				fetchWait('');
//				$('#p_sagyo_log_update').text('□');
				var contentType = res.headers.get("content-type");
				debug_toast('<div class="h7">p_sagyo_log:' + contentType + '</div>');
				if(contentType && contentType.indexOf("application/json") !== -1) {
					return res.json();
				} else {
//					$('#p_sagyo_log').find('.tm').text('Error');
					$('#p_sagyo_log_update').text('Error');
					$.toast({
						text : 'p_sagyo_log error:' + contentType + ' ' + url
						,loader: false
						,hideAfter : 60000
					});
				}
			} )
			.then((json) => {
				debug_toast('<div class="h7">p_sagyo_log:json:' + json.data.length + '</div>');
				var td = {};	//[];	//{};
				var	tanto = '';
				var	nin = 0;
				for ( var i = 0 ; i < json.data.length ; i++ ) {
					if(tanto != json.data[i].TANTO_CODE) {
						loc = json.data[i].Loc;
						if(!td[loc]) {
							td[loc] = '';
						}
						td[loc] += '<span>' + json.data[i].TANTO_NAME + '</span>';
						tanto = json.data[i].TANTO_CODE;
						nin++;
					}
				}
	/*
				td.sort(function(a, b) {
					var a_key = a;
					var b_key = b;
					if (a_key < b_key) return -1;
					if (a_key > b_key) return 1;
					return 0;
				});
	*/
				var	keys = [];
				for (var key in td) {
					keys.push(key);
				}
				keys.sort();
				var	tr = '';
				for(v = 0;v < keys.length;v++) {
					tr += '<tr>';
					tr += '<td>' + keys[v] + '</td>';
					tr += '<td>' + td[keys[v]] + '</td>';
					tr += '</tr>';
				}
				tr += '<tr>';
				tr += '<td colspan="2">作業人数：'+ nin + ' 名(※30分以内)</td>';
				tr += '</tr>';
	/*
				for (var t in td) {
					if(t != '') {
						loc_name = t;
						tr += '<tr>';
						tr += '<td>' + t + '</td>';
						tr += '<td>' + td[t] + '</td>';
						tr += '</tr>';
					}
				}
	*/
				if($(window).scrollTop() == 0) {
//2018.12.28		$('#p_sagyo_log > table').floatThead('destroy');
				}
				$('#p_sagyo_log > table > tbody').find("tr").remove();
				$('#p_sagyo_log > table > tbody').append(tr);
//				$('#p_sagyo_log').find('.tm').text(datetimeFormat(now).slice(-5));
//				$('#p_sagyo_log_update').text(nowTM());
				$(this).text('□更新.' + nowTM());
				$(this).offset({
					 top: $('#p_sagyo_log > table').offset().top - $(this).height()
					,left: $('#p_sagyo_log > table').offset().left + $('#p_sagyo_log > table').width() - $(this).width()
				});
//				$('#p_sagyo_log > table').floatThead();
			} );
		}
		if($('#p_sagyo_log_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					debug_toast('setTimeout:#p_sagyo_log:' + $('#navId').text());
					$('#p_sagyo_log_update').trigger("click");
				},$('#p_sagyo_log_fetch').val() * 1000);
			}
		}
		return false;
	});
	$('a[href~="#p_sagyo_log"]').text('作業状況');
	if($('#p_sagyo_log_chg').val() > 0) {
		$('#p_sagyo_log_update').trigger("click");
		$('#p_sagyo_log').removeClass('disable');
	} else {
		$('#p_sagyo_log').addClass('disable');
	}
});
//keyでソートする
function keySort(hash,sort){
//	var sortFunc = sort;
	var keys = [];
	var newHash = {};
	for (var k in hash) keys.push(k);
	keys.sort();
	var length = keys.length;
	for(var i = 0; i < length; i++){
		newHash[keys[i]] = hash[keys[i]];
	}
	return newHash;	
}
/*
エアコン発注残
*/
$(document).ready(function() {
	setConfig('#AcOrder_chg','');
	setConfig('#AcOrder_scr','');
	setConfig('#AcOrder_fetch',600);
	$('#AcOrder').on('focus', function() {
		$('#navChg').text($('#AcOrder_chg').val());
		$('#navScr').text($('#AcOrder_scr').val());
		$('#AcOrder_update').offset({
			 top: $('#AcOrder > table').offset().top - $('#AcOrder_update').height()
			,left: $('#AcOrder > table').offset().left + $('#AcOrder > table').width() - $('#AcOrder_update').width()
		});
	});
	var timer = null;
	$('#AcOrder_update').on('click', function(){
		$(this).text($(this).text().replace('□','■'));
		// 同時処理防止
		if(fetchWait(this.id) == false) {
			debug_toast('click:' + this.id);
			var	url = 'monitor.py?f=AcOrder&dns=' + $('#dns').val();
			//検索実行
			debug_toast('<div class="h7">' + url + '</div>');
			fetch(url)
			.then((res) => {
				// 同時処理解除
				fetchWait('');
//				$('#AcOrder_update').text('□');
				var contentType = res.headers.get("content-type");
				debug_toast('<div class="h7">AcOrder:' + contentType + '</div>');
				if(contentType && contentType.indexOf("application/json") !== -1) {
					return res.json();
				} else {
					$.toast({
						text : 'AcOrder error:' + contentType + ' ' + url
						,loader: false
						,hideAfter : 60000
					});
				}
			} )
			.then((json) => {
				debug_toast('<div class="h7">AcOrder:json:' + json.data.length + '</div>');
				var	th = '<tr><th></th>';
				var	td1 = '<tr><th>件数</th>';
				var	td2 = '<tr><th class="zan0">残</th>';
				var	td3 = '<tr><th>個数</th>';
				var	td4 = '<tr><th class="zan0">残</th>';
				var	col = json.data.length;
				if (col > 10) {
					col = 10;
				}
				if($(window).scrollTop() == 0) {
//2018.12.28		$('#AcOrder > table').floatThead('destroy');
				}
				$('#AcOrder > table > thead').find("tr").remove();
				$('#AcOrder > table > tbody').find("tr").remove();
				for ( var i = 0 ; i < json.data.length ; i++ ) {
					if((i % 10) == 0) {
						if(i > 0) {
							th += '</tr>';
							td1 += '</tr>';
							td2 += '</tr>';
							td3 += '</tr>';
							td4 += '</tr>';
							$('#AcOrder > table > thead').append(th);
							$('#AcOrder > table > tbody').append(td1 + td2 + td3 + td4);
							th = '<tr><td colspan="11"></td></tr><tr class="date"><th></th>';
						} else {
//							th = '<tr class="date"><th>指定納期</th>';
							th = '<tr class="date"><th>指定納期</th>';
						}
						td1 = '<tr class="cnt"><th>商品化 件数</th>';
						td2 = '<tr><th class="zan0">商品化 件数残</th>';
						td3 = '<tr class="qty"><th>商品化 個数</th>';
						td4 = '<tr><th class="zan0">商品化 個数残</th>';
					}
					var	dt = json.data[i].NokiDsp;
					if(isNumber(dt.slice(0,2))) {
						var	dt_text = Number(dt.slice(0,2)) + '/' + Number(dt.slice(-2));
//						if(i > 0) {
//							var	prev = json.data[i-1].NokiDsp;
//							if(Number(dt.slice(0,2)) == Number(prev.slice(0,2))) {
//								var	dt_text = Number(dt.slice(-2));
//							}
//						}
						th += '<td tabindex="0">' + dt_text + '</td>';
					} else {
						th += '<td class="h7 normal" tabindex="0">' + dt + '</td>';
					}
					td1 += '<td>' + json.data[i].sdcCnt + '</td>'
					td2 += '<td>' + json.data[i].zanCnt + '</td>'
					td3 += '<td>' + json.data[i].sdcQty + '</td>'
					td4 += '<td>' + json.data[i].zanQty + '</td>'
				}
				if((i % 10) != 0) {
					th += '</tr>';
					td1 += '</tr>';
					td2 += '</tr>';
					td3 += '</tr>';
					td4 += '</tr>';
					$('#AcOrder > table > thead').append(th);
					$('#AcOrder > table > tbody').append(td1 + td2 + td3 + td4);
				}
				$('#AcOrder > table > tbody').change();
//				console.log('position().left:' + $('#AcOrder > table').position().left);
//				console.log('offset().left:' + $('#AcOrder > table').offset().left);
				console.log(this.id + '.offset().left:' + $(this).offset().left);
				$(this).text('□更新.' + nowTM());
				$(this).offset({
					 top: $('#AcOrder > table').offset().top - $(this).height()
					,left: $('#AcOrder > table').offset().left + $('#AcOrder > table').width() - $(this).width()
				});
//				$('#AcOrder > table').floatThead();
			} );
		}
		if($('#AcOrder_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					debug_toast('setTimeout:#AcOrder:' + $('#navId').text());
					$('#AcOrder_update').trigger("click");
				},$('#AcOrder_fetch').val() * 1000);
			}
		}
		return false;
	});
	$('#AcOrder > table > tbody').change(function(){
		for(row = 0;row < $('tr',this).length;row += 6) {
			for(i = 0;i < $('td',$('tr',this).eq(row)).length;i++) {
				var	zan = $('td',$('tr',this).eq(row + 1)).eq(i);
				var v = $('td',$('tr',this).eq(row)).eq(i);
				var	r = $(zan).text() / $(v).text();
				var	clr = 'rgb(0,0,' + parseInt(139 * r) + ')';
				$(zan).css({
					backgroundColor: clr
				});
				var	zan = $('td',$('tr',this).eq(row + 3)).eq(i);
				var v = $('td',$('tr',this).eq(row + 2)).eq(i);
				var	r = $(zan).text() / $(v).text();
				var	clr = 'rgb(0,' + parseInt(100 * r) + ',0)';
				$(zan).css({
					backgroundColor: clr
				});
			}
		}
	});
	$('a[href~="#AcOrder"]').text('エアコン発注残');
	if($('#AcOrder_chg').val() > 0) {
		$('#AcOrder_update').trigger("click");
		$('#AcOrder').removeClass('disable');
	} else {
		$('#AcOrder').addClass('disable');
	}
});
/*
仮置き状況
*/
var	zaiko9_data = [];
$(document).ready(function() {
	setConfig('#zaiko9_chg',10);
	setConfig('#zaiko9_scr','');
	setConfig('#zaiko9_fetch',600);
	$('#zaiko9').on('focus', function() {
		$('#navChg').text($('#zaiko9_chg').val());
		$('#navScr').text($('#zaiko9_scr').val());
		utimeOffset('#zaiko9_update','#' + this.id);
/*
		console.log(this.id + '.focus()');
		console.log('#zaiko9_update.offset().top: ' + $('#zaiko9_update').offset().top);
		console.log('#zaiko9_update.offset().left: ' + $('#zaiko9_update').offset().left);
		$('#zaiko9_update').offset({
			 top: $(window).scrollTop() + $('#zaiko9').offset().top - $('#zaiko9_update').height()
			,left: $('#zaiko9 > table').offset().left + $('#zaiko9 > table').width() - $('#zaiko9_update').width()
		});
*/
	});
	var timer = null;
	$('#zaiko9_update').on('click', function(){
//		if($('#navId').text() == '#zaiko9' || $(this).text() == '◇') {
		$(this).text($(this).text().replace('□','■'));
		// 同時処理防止
		if(fetchWait(this.id) == false) {
			debug_toast('click:' + this.id);
			var	url = 'monitor.py?f=zaiko9&dns=' + $('#dns').val();
			if($('#JGYOBU').val() != '') {
				url += ('&JGYOBU=' + $('#JGYOBU').val());
			}
			//検索実行
			debug_toast('<div class="h7">' + url + '</div>');
			fetch(url)
			.then((res) => {
				// 同時処理解除
				fetchWait('');
				var contentType = res.headers.get("content-type");
				debug_toast('<div class="h7">zaiko9:' + contentType + '</div>');
				if(contentType && contentType.indexOf("application/json") !== -1) {
					return res.json();
				} else {
					$.toast({
						text : 'AcOrder error:' + contentType + ' ' + url
						,loader: false
						,hideAfter : 60000
					});
				}
			} )
			.then((json) => {
				debug_toast('<div class="h7">zaiko9:json:' + json.data.length + '</div>');
				zaiko9_data = json.data;
				var	tr = '';
				var	soko = ''
				for ( var i = 0 ; i < json.data.length ; i++ ) {
					debug_toast('<div class="h7">zaiko9:' + json.data[i].Pn + '</div>');
					if(json.data[i].Soko != soko) {
						soko = json.data[i].Soko;
						tr += '<tr><td class="soko" colspan="5">' + soko + ' ' + json.data[i].SOKO_NAME  + '</td><tr>';
					}
					tr += '<tr>';
					tr += '<td>' + json.data[i].Tana + '</td>';
					tr += '<td>' + json.data[i].Pn + '</td>';
					tr += '<td>' + json.data[i].Qty + '</td>';
					tr += '<td>' + json.data[i].inDate + '</td>';
					tr += '<td>' + json.data[i].stTana + '</td>';
					tr += '</tr>';
				}
				if($(window).scrollTop() == 0) {
//2018.12.28		$('#zaiko9 > table').floatThead('destroy');
				}
				$('#zaiko9 > table > tbody').find("tr").remove();
//				console.log(tr);
				$('#zaiko9 > table > tbody').append(tr);
				$(this).text('□更新.' + nowTM());
				utimeOffset('#' + this.id, '#zaiko9');

				if($('#y_syuka_h_chg').val() > 0) {
					if(zaiko9_data.length > 0) {
						var	qty90 = 0;
						for ( var i = 0 ; i < zaiko9_data.length ; i++ ) {
							if(zaiko9_data[i].Soko == '90') {
								qty90++;
							}
						}
						$.toast({
							text : '<div>仮置き残(90)：' + qty90 + '件</div>'
							,textAlign : 'left'
							,bgColor : 'yellow'
							,textColor : 'black'
							,position : 'bottom-left'
							,loader: false
							,hideAfter : $('#zaiko9_fetch').val() * 1000
						});
					}
					return ;
				}

			} );
		}
		if($('#zaiko9_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					debug_toast('setTimeout:#zaiko9:' + $('#navId').text());
					$('#zaiko9_update').trigger("click");
				},$('#zaiko9_fetch').val() * 1000);
			}
		}
		return false;
	});
//	$('#zaiko9 > table > tbody').change(function(){
//	});
//	$('a[href~="#zaiko9"]').text('入庫待ち在庫');
	if($('#zaiko9_chg').val() > 0) {
		$('#zaiko9_update').trigger("click");
		$('#zaiko9').removeClass('disable');
	} else {
		$('#zaiko9').addClass('disable');
	}
});

var	scroll_ID = "";
$(window).scroll(function() {
//	console.log("scroll:" + this.id);
//	console.log("scroll:" + $(this).scrollTop() + " " + $('#navId').text() + ":" + $($('#navId').text()).offset().top);
	if(scroll_ID != $('#navId').text()) {
		scroll_ID = $('#navId').text();
//		$('table', scroll_ID).floatThead({top: $(scroll_ID).offset().top});
//		$('table',this).floatThead('reflow');
	}
});
/*
画面切替
*/
$(document).ready(function() {
	$('.content').on('focus', function(){
		var title = $('a[href~="#' + this.id + '"]').text();
		console.log('.content.focus:' + this.id + ':' + title);
//		debug_toast('focus:' + this.id + ' ' + title);
		$('#title').text(title);
//		$('#navCnt').text('0');
//		$("html").animate({ scrollTop: 0});
//		var	top = $(this).offset().top;
//		debug_toast(this.id + ' floatThead top:' + top);
//		$('table').floatThead('destroy');
/* temp
		$('table',this).floatThead({
							top: top,
							position: 'auto',	//'auto','fixed','absolute'
							});
		$('table',this).floatThead('reflow');
*/
//		var	top = $('header').offset().top;
//		debug_toast('floatThead:header:' + top);
//		$('header').floatThead({
//							top: top
//							});
//		return false;
	});
/*
	$('.content').on('blur', function(){
		$('table',this).floatThead('destroy');
		return false;
	});
*/
/*
	$('.content').on('click', function(){
		console.log('click:' + this.id);
		console.log('$(#navId).text():' + $('#navId').text());
//		$('#navCnt').text(0);
//		$('#navChg').text(30);
		if($('#navId').text() != '#config_div') {
			debug_toast('click:' + this.id);
//			$('nav').navChange(false);
			return false;
		}
	});
*/
	$('.content').on('click', function(e){
		console.log('.content click:' + this.id + ' ' + typeof e.which);
//		if(typeof e.which !== 'undefined') {
		if(e.which) {
//			console.log($('#auto').text() + ': stop');
			$.toast().reset('all');
			if($('#play').text() != 'pause') {
				$('#play').text('pause');
				$.toast({text: '自動切替：OFF',	loader: false});
			}
//			return false;
		}
	});
});
/*
画面スクロール
*/
$(document).ready(function() {
	$('#title').on('click', function(){
//		$("html").stop();
//		$($('#navId').text()).trigger("focus");
		if($(window).scrollTop() != 0) {
	        $("html").animate({ scrollTop: 0});
//			$(window).scrollTop(0);
			return false;
		}
        var end_pos = $('footer').offset().top;
		debug_toast('<div class="h7">スクロール開始↓:' + end_pos + '</div>');
        $("html").animate(
			{ scrollTop: end_pos },
			{
				duration: end_pos * 20,	//50 * 1000, 	//'slow',
				easing: 'linear',		//'swing',
				complete: function() {
					debug_toast('<div class="h7">スクロール開始↑:0</div>');
			        $(this).animate(
						{ scrollTop: 0 },
						{
							duration: end_pos * 10,	//50 * 1000, 	//'slow',
							easing: 'linear',		//'swing',
							complete: function() {
								debug_toast('<div class="h7">スクロール終了</div>');
							}
						}
					);
				}
            }
		);
		return false;
	});
});

/*
商品化状況(JCS) p_sshiji_jcs
*/
$(document).ready(function() {
	$('#p_sshiji_jcs').on('focus', function() {
		utimeOffset('#p_sshiji_jcs_update','#' + this.id);
	});
	$('#p_sshiji_jcs_update').on('click', function() {
		console.log(this.id + '.click()');
		$(this).text($(this).text().replace('□','■'));
		if($(window).scrollTop() == 0) {
//2018.12.28	$('#p_sshiji_jcs > table').floatThead('destroy');
		}
		$('#p_sshiji_jcs > table > tbody').find("tr").remove();
		for( var i = 0; $('#p_sshiji_jcs > table > thead > tr:first th').length > 1; i++) {
			console.log("th.length:" + $('#p_sshiji_jcs > table > thead > tr:first th').length);
			$('#p_sshiji_jcs > table > thead > tr:first th:last').remove();
		}

//		var	p_sshiji_tr = $('#p_sshiji > tbody > tr');
		var	row = 0;
		var	col = 0;
		var	noki_cur = "-";
//		["事前","ＳＸ","当日出荷分","出荷分"].forEach(function( value ) {
		["0","2","1"].forEach(function( value ) {
			for( i = 0; i < p_sshiji_data.length; i++) {
				row++;
				var	cls = '';
//				var	kubun = $('td',p_sshiji_tr.eq(i)).eq(3).text();
				var	shiji_f = p_sshiji_data[i].SHIJI_F;
				if(value == shiji_f) {
					switch(shiji_f) {
					case '0':		cls = 'jizen';	break;//'事前'
//					case 'ＳＸ':	cls = 'jizen';	break;//'ＳＸ'
					case '1':		cls = 'spot';	break;//'出荷分'
					case '2':		cls = 'keppin';	break;//'当日出荷分'
//					case '再梱包':						break;
					}
//					var	noki_new = $('td',p_sshiji_tr.eq(i)).eq(2).text();
					var	noki_new = p_sshiji_data[i].AcNoki;
					if(noki_cur != noki_new) {
						col++;
						var	th = "";
						var	dt = noki_new;
						if(dt != '') {
							if(dt.search( /\d{4}\-\d{2}\-\d{2}/ ) == 0) {
								result = dt.match( /(\d{4})\-(\d{2})\-(\d{2})/ );
								dt = Number(result[2]) + '/' + Number(result[3]);
							} else if(dt.search( /\d{8}/ ) == 0) {
								var	dt = noki_new.slice(-4);
								dt = Number(dt.slice(0,2)) + '/' + Number(dt.slice(-2));
							}
						} else {
							dt = 'ＳＸ';
						}
						th += '<th class="' + cls + '" title="' + cls + '" colspan="2">' + dt + '</th>';
						$('#p_sshiji_jcs > table > thead > tr:first').append(th);
						noki_cur = noki_new;
						row = 1;
					}
//					if($('td',p_sshiji_tr.eq(i)).eq(7).text() != "") {
					if(kanCheck(p_sshiji_data[i]) != 0) {
						cls = 'sumi';
					}
					var	td = "";
//					td += '<td class="' + cls + '">' + $('td',p_sshiji_tr.eq(i)).eq(5).text() + '</td>';
					td += '<td class="' + cls + '">' + p_sshiji_data[i].HIN_GAI + '</td>';
//					td += '<td class="' + cls + '">' + $('td',p_sshiji_tr.eq(i)).eq(6).text() +  '</td>';
					td += '<td class="' + cls + '">' + p_sshiji_data[i].qty +  '</td>';

					if(row > $('#p_sshiji_jcs > table > tbody > tr').length) {
						var	tr = "";
						tr += "<tr>";
						tr += "<td>" + row + "</td>";
//						tr += td;
						tr += "</tr>";
						$('#p_sshiji_jcs > table > tbody').append(tr);
					}
					var	tr_col = ($('td',$('#p_sshiji_jcs > table > tbody > tr').eq(row - 1)).length - 1) / 2 + 1;
					for( ; tr_col < col ; tr_col++) {
						td = '<td class="none"></td><td class="none"></td>' + td
					}
					$('#p_sshiji_jcs > table > tbody > tr').eq(row - 1).append(td);
				}
			}
		});
		$(this).text('□更新.' + nowTM());
		utimeOffset('#' + this.id, '#p_sshiji_jcs');
//		$('#p_sshiji_jcs > table').floatThead({top: $('#p_sshiji_jcs > table').offset().top});
	});
/*
		var	tr = "";
		var	row = 0;
		var	list = [];
		var	p_sshiji_tr = $('#p_sshiji > tbody > tr');
			var	recs = {};
			recs["Pn"] = $('td',p_sshiji_tr.eq(i)).eq(5).text();
			recs["Qty"] = $('td',p_sshiji_tr.eq(i)).eq(6).text();
			list.push(recs);
		}
			row++;
			tr += "<tr>";
			tr += "<td>" + row + "</td>";
			tr += "<td>" + $('td',p_sshiji_tr.eq(i)).eq(5).text() + "</td>";
			tr += "<td>" + $('td',p_sshiji_tr.eq(i)).eq(6).text() + "</td>";
			tr += "</tr>";
		$('#p_sshiji_jcs > table > tbody').append(tr);
*/
});
/*
商品化状況(区分) p_sshiji_cat
*/
$(document).ready(function() {
	$('.jump_to_top').on('click', function() {
		$('html').scrollTop(0);
	});

	setConfig('#p_sshiji_cat_chg','');
	setConfig('#p_sshiji_cat_scr','');
	setConfig('#p_sshiji_cat_fetch','');
	$('#p_sshiji_cat').on('focus', function() {
		utimeOffset('#p_sshiji_cat_update','#' + this.id);
		$('#navChg').text($('#p_sshiji_cat_chg').val());
		$('#navScr').text($('#p_sshiji_cat_scr').val());
	});
	if($('#p_sshiji_cat_chg').val() > 0) {
		$('#p_sshiji_cat').removeClass('disable');
	} else {
		$('#p_sshiji_cat').addClass('disable');
	}
	var	dt = [];
	$('#p_sshiji_cat_update').on('click', function() {
		console.log(this.id + '.click()');
		$(this).text($(this).text().replace('□','■'));
		if($(window).scrollTop() == 0) {
//2018.12.28	$('#p_sshiji_cat > table').floatThead('destroy');
		}
		//ヘッダー行更新
		var	th = $('th', '#p_sshiji_cat > table > thead > tr:first');
//		$(th).text('');
//		$(th).eq(0).text('区分');
		var	tbody = $('tr', '#p_sshiji_cat > table > tbody');
		$(tbody).find('td').text('');
		$('td', $(tbody).eq(0)).eq(0).text('コンデンサ');
		$('td', $(tbody).eq(1)).eq(0).text('フィルター');
		$('td', $(tbody).eq(2)).eq(0).text('パイプ');
		$('td', $(tbody).eq(3)).eq(0).text('モータファン');
		$('td', $(tbody).eq(4)).eq(0).text('エバポレータ');
		$('td', $(tbody).eq(5)).eq(0).text('ヒータコア');
		$('td', $(tbody).eq(6)).eq(0).text('Oリング');
		$('td', $(tbody).eq(7)).eq(0).text('その他');
		$('td', $(tbody).eq(8)).eq(0).text('※NG=Z(在庫なし)を集計');
		var	col = 0;
		dt = [''];
		for(var i = 0; i < jcs_list.length; i++) {
			if(dt[dt.length - 1] != jcs_list[i].DlvDt) {
				col++;
				dt.push(jcs_list[i].DlvDt);
				var	ymd = jcs_list[i].DlvDt.replace(/-/g,'');
				var	md = ymd.slice(-4);
				md = Number(md.slice(0,2)) + '/' + Number(md.slice(-2));
				$(th).eq(col).text(md);
				$(th).eq(col).attr('title', ymd);
			}
			var	row = 2;
			switch(jcs_list[i].CATEGORY_CODE) {
			case "C01":	row = 0;	break;
			case "C02":	row = 1;	break;
			case "C03":	row = 2;	break;
			case "C04":	row = 3;	break;
			case "C05":	row = 4;	break;
			case "C06":	row = 5;	break;
			case "C07":	row = 6;	break;
			case "C99":	row = 7;	break;
			default:	row = 7;	break;
			}
			var	td = $('td', $(tbody).eq(row)).eq(col);
			var	qty = Number($(td).text());
//			if(jcs_list[i].CANCEL_F != '1' && jcs_list[i].SHONIN_CODE == '') {
			if(jcs_list[i].CANCEL_F != '1' && (jcs_list[i].SHONIN_CODE == '' || jcs_list[i].SHONIN_CODE == 'Z')) {
				qty += Number(jcs_list[i].Qty);
			}
			$(td).text(qty);
		}
		$(this).text('□更新.' + nowTM());
		utimeOffset('#' + this.id, '#p_sshiji_cat');
		//list
		jcs_list.sort(function(a, b) {
			if (a.DlvDt > b.DlvDt) {
				return 1;
			} else if (a.DlvDt < b.DlvDt) {
				return -1;
			}
			if (a.CATEGORY_CODE > b.CATEGORY_CODE) {
				return 1;
			} else if (a.CATEGORY_CODE < b.CATEGORY_CODE) {
				return -1;
			}
			if (a.PACKING_NO > b.PACKING_NO) {
				return 1;
			} else if (a.PACKING_NO < b.PACKING_NO) {
				return -1;
			}
			if (a.K_KEITAI > b.K_KEITAI) {
				return 1;
			} else {
				return -1;
			}
		});
		var	tr = '';
		var	curr = '';
		var	row = 0;
		var	day = 0;
		for(var i = 0; i < jcs_list.length; i++) {
/*
			if(curr != jcs_list[i].DlvDt) {
				curr = jcs_list[i].DlvDt;
				day++;
				row = 0;
				tr += '<tr>';
				tr += '<td class="row"></td>';
				tr += '<td colspan="7">';
				tr += '<a href="#day' + (day - 1) + '">▲</a>';
				tr += '<a href="#day' + (day + 1) + '" id="day' + day + '">▼</a>';
				tr += '</td>';
				tr += '</tr>';
			}
*/
			row++;
			tr += '<tr>';
			tr += '<td class="row">';
			tr += row;
			tr += '</td>';
			tr += '<td class="noki">';
			tr += jcs_list[i].DlvDt;
			tr += '</td>';
			tr += '<td class="pn">';
			tr += jcs_list[i].MazdaPn;
			tr += '</td>';
			tr += '<td class="name"><div>';
			tr += jcs_list[i].NameJ != '' ? jcs_list[i].NameJ : jcs_list[i].NameE;
			tr += '</div></td>';
			tr += '<td class="qty">';
			tr += jcs_list[i].Qty;
			tr += '</td>';
			tr += '<td class="kubun ' + jcs_list[i].CATEGORY_CODE + '">';
			var	nm = 'その他';
			switch(jcs_list[i].CATEGORY_CODE) {
			case 'C01':	nm = 'コンデンサ';		break;
			case 'C02':	nm = 'フィルター';		break;
			case 'C03':	nm = 'パイプ';			break;
			case 'C04':	nm = 'モータファン';	break;
			case 'C05':	nm = 'エバポレータ';	break;
			case 'C06':	nm = 'ヒータコア';		break;
			case 'C07':	nm = 'Oリング';			break;
			}
			tr += nm;
			tr += '</td>';
			tr += '<td class="spn ' + jcs_list[i].CATEGORY_CODE + '">';
			tr += jcs_list[i].PACKING_NO;
			tr += '</td>';
			tr += '<td class="type ' + jcs_list[i].CATEGORY_CODE + '">';
			tr += jcs_list[i].K_KEITAI;
			tr += '</td>';
			tr += '<td class="stat">';
			if(jcs_list[i].CANCEL_F == '1') {
				tr += 'ｷｬﾝｾﾙ';
			} else if(jcs_list[i].SHONIN_CODE != '') {
				tr += jcs_list[i].SHONIN_CODE;
			}
			tr += '</td>';
			tr += '</tr>';
		}
		$('#p_sshiji_cat > table.list > tbody').find("tr").remove();
		$('#p_sshiji_cat > table.list > tbody').append(tr);
	});
	$(document).on('*mouseenter', '#p_sshiji_cat > table > tbody > tr > td', function(){
		console.log('td.mouseenter ' + $(this).text());
//		var	th = $('th', '#p_sshiji_cat > table > thead > tr');
//		var	th = $('#p_sshiji_cat > table > thead > tr').find('th');
/*
		var	th = $('#p_sshiji_cat > table > tbody > tr').find('td');
		console.log('th.length = ' + th.length);
		for(var i = 0; i < th.length; i++) {
			console.log('th[' + i + '].text()=' + $(th).eq(i).text());
		}
*/
		$tag_td = $(this)[0];
		$tag_tr = $(this).parent()[0];
		var	txt = '';
		txt += '<table class="width-100">';
		txt += '<tr><td colspan="3">';
		txt += '' + dt[$tag_td.cellIndex] + '';
		txt += ' ' + $('td', $tag_tr).eq(0).text() + '';
//		txt += ' (' + $tag_tr.rowIndex + ',' + $tag_td.cellIndex + ')';
		txt += '</td><td class="text-right" colspan="2">' + $(this).text() + '</td>';
		txt += '</tr>';

		for(var i = 0; i < jcs_list.length; i++) {
			if(dt[$tag_td.cellIndex] == jcs_list[i].DlvDt) {
				var	scat = jcs_list[i].CATEGORY_CODE;
				if(scat == '') {
					scat = '小物';
				}
				if($('td', $tag_tr).eq(0).text() == scat) {
					txt += '<tr>';
					txt += '<td>' + jcs_list[i].SHIJI_NO + '</td>';
					txt += '<td>' + jcs_list[i].Pn + '</td>';
					txt += '<td>' + jcs_list[i].NameJ + '</td>';
					txt += '<td>' + jcs_list[i].PACKING_NO + '</td>';
					txt += '<td class="text-right">' + jcs_list[i].Qty + '</td>';
					txt += '</tr>';
				}
			}
		}
		txt += '</table>';
		$('#dialog').html(txt);
		$('#dialog').show();
		console.log('offset=' + $(this).offset());
		var	left = $(this).offset().left;
		if((left + $('#dialog').width()) > $(window).width()) {
			left = $(window).width() - $('#dialog').width() - 10;
		}
		$('#dialog').offset({
							 top: $(this).offset().top + $(this).height(),
							 left: left
							 });
	});
	$(document).on('*mouseleave', '#p_sshiji_cat > table > tbody > tr > td', function(){
		console.log('td.mouseleave ' + $(this).text());
		$('#dialog').hide();
	});
});
/*
勤務予定 work_sch
*/
$(document).ready(function() {
	setConfig('#work_sch_chg','');
	setConfig('#work_sch_scr','');
	setConfig('#work_sch_fetch','');
	setConfig('#work_sch_mode','');
	$('#work_sch').on('focus', function() {
		utimeOffset('#work_sch_update','#' + this.id);
		$('#navChg').text($('#work_sch_chg').val());
		$('#navScr').text($('#work_sch_scr').val());
	});
/*
	$('#work_sch').on('click', function() {
		console.log('#' + this.id + '.click()');
		$('#navChg').text('100');
		$('#navCnt').text('0');
	});
*/
	$('#work_sch > table > thead > tr > .prev').on('click', function() {
//		alert('next');
		console.log(this.id + '.click() prev');
		var	dt = new Date($(this).attr('title'));
		dt.setDate(dt.getDate() - 7);
		//alert(dateYMD(prev,'-'));
		$('#work_sch_update').attr('title', dateYMD(dt,'-'));
		$('#work_sch_update').trigger('click');
	});
	$('#work_sch > table > thead > tr > .next').on('click', function() {
//		alert('next');
		console.log(this.id + '.click() next');
		var	dt = new Date($(this).attr('title'));
		dt.setDate(dt.getDate() + 1);
		$('#work_sch_update').attr('title', dateYMD(dt,'-'));
		$('#work_sch_update').trigger('click');
	});
	var timer = null;
	$('#work_sch_update').on('click', function() {
		console.log(this.id + '.click()');
		$(this).text($(this).text().replace('□','■'));
		var	url = 'sch.py?dns=' + $('#dns').val();
		if($('#work_sch_mode').val() == '') {
			url += '&post=01';
		}
		var title = $('#work_sch_update').attr('title');
		if(typeof title !== 'undefined' && title != '') {
			url += '&stdate=' + title;
			$('#work_sch_update').attr('title','');
		}
		fetch(url)
		.then((res) => {
			// 同時処理解除
			fetchWait('');
			return res.json();
		})
		.then((json) => {
//			$('#work_sch > table').floatThead('destroy');
			var	th = $('th', '#work_sch > table > thead > tr:first');
			var	tanto_code = '';
			var	tr = '';
			for ( var i = 0 ; i < json.data.length ; i++ ) {
				if(tanto_code != json.data[i].TANTO_CODE) {
					if(tanto_code != '') {
						tr += '</tr>';
					}
					tanto_code = json.data[i].TANTO_CODE;
					var title = tanto_code;
					title += ' ' + json.data[i].TANTO_NAME;
					tr += '<tr class="click ' + $('#work_sch_mode').val() + '"><td title="' + title + '">';
					tr += '<div class="tanto_code">' + tanto_code + '</div>';
					tr += '<div class="tanto_name">' + json.data[i].SurName + '</div></td>';
				}
				var title = json.data[i].CalDate;
				title += ' ' + json.data[i].sTANTO_CODE;
				tr += '<td title="' + title + '">';
				var	div = '';
				if($('#work_sch_mode').val() != 'w5') {
					div += '<div class="btn off holiday">休日</div>';
					div += '<div class="btn off filter">フィルター</div>';
					div += '<div class="btn off condenser">コンデンサ</div>';
					div += '<div class="btn off pipe">パイプ</div>';
					div += '<div class="btn off petty">小物</div>';
				} else {
					div += '<div class="btn off am">AM</div>';
					div += '<div class="btn off pm">PM</div>';
				}
				if(json.data[i].WorkDet.indexOf('休日') != -1) {
					div = div.replace('off holiday','holiday');
				}
				if(json.data[i].WorkDet.indexOf('フィルター') != -1) {
					div = div.replace('off filter','filter');
				}
				if(json.data[i].WorkDet.indexOf('コンデンサ') != -1) {
					div = div.replace('off condenser','condenser');
				}
				if(json.data[i].WorkDet.indexOf('パイプ') != -1) {
					div = div.replace('off pipe','pipe');
				}
				if(json.data[i].WorkDet.indexOf('小物') != -1) {
					div = div.replace('off petty','petty');
				}
				if(json.data[i].WorkDet.indexOf('AM') != -1) {
					div = div.replace('off am','am');
				}
				if(json.data[i].WorkDet.indexOf('PM') != -1) {
					div = div.replace('off pm','pm');
				}
				tr += div + '</td>';
				if( i < 7 ) {
					var	dt = new Date(json.data[i].CalDate);
					var	html = '<div class="date">' + (dt.getMonth() + 1) + '/' + dt.getDate() + '</div>';
					html += '<div class="dow' + dt.getDay() + '">' + '日月火水木金土'[dt.getDay()] + '</div>';
					console.log(i + ':' + $(th).eq(i + 1).html());
					console.log('→' + html);
					$(th).eq(i + 1).html(html);
					$(th).eq(i + 1).attr('title', json.data[i].CalDate);
				}
			}
			if(tanto_code != '') {
				tr += '</tr>';
			}
			$('#work_sch > table > tbody').find("tr").remove();
			$('#work_sch > table > tbody').append(tr);
			$(this).text('□更新.' + nowTM());
			utimeOffset('#' + this.id, '#work_sch');
			//予定 編集ダイアログ
			$('#work_sch > table > tbody > tr').on('click', function() {
//				$('#work_sch > table').floatThead('destroy');
//				console.log('tr.click:' + $(this).html());
				console.log('tr.click:' + $(this)[0].rowIndex);
				var	html = '<table>';
				html += '<caption>' + $(this)[0].rowIndex + '</caption>';
				html += $('#work_sch > table > thead').html();
				html += '<tbody><tr>';
				for(i = 0;i < $('td',this).length; i++) {
					html += '<td title="' + $('td',this).eq(i).attr('title') + '">';
//					if(i == 0) {
//						html += $('td',this).eq(i).html();
//					} else {
					if($('td',this).eq(i).html() != '') {
						html += $('td',this).eq(i).html();
					} else {
						html += '<div class="btn off holiday">休日</div>';
						html += '<div class="btn off filter">フィルター</div>';
						html += '<div class="btn off condenser">コンデンサ</div>';
						html += '<div class="btn off pipe">パイプ</div>';
						html += '<div class="btn off petty">小物</div>';
					}
//					}
					html += '</td>';
				}
				html += '</tr></tbody>';
				html += '</table>';
				console.log(html);
				$('#work_sch_edit').html(html);
		        $("#work_sch_edit").dialog("open");
				$('#work_sch_edit .btn').on('click', function() {
					console.log('click->' + $(this).text);
					if($(this).hasClass('off')) {
						$(this).removeClass('off');
					} else {
						$(this).addClass('off');
					}
				});
			});
		})
		.catch((err) => {
			console.log(err);
			$.toast({text: '<div class="h7">' + url + '</div><div class="h7">fetch err:' + err + '</div>',loader: false, hideAfter : 60000});
		});
		if($('#work_sch_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					$('#work_sch_update').trigger("click");
				},$('#work_sch_fetch').val() * 1000);
			}
		}
		return false;
	});
	if($('#work_sch_chg').val() > 0) {
		$('#work_sch_update').trigger("click");
		$('#work_sch').removeClass('disable');
	} else {
		$('#work_sch').addClass('disable');
	}
/*
	if($('#UKEHARAI_CODE').val() == 'ZI0') {
		$('#work_sch').removeClass('disable');
		$('#work_sch_update').trigger("click");
	}
*/

    $("#work_sch_edit").dialog({
        autoOpen: false,
        height: 'auto',
        width: 'auto',
        modal: true,
        buttons: {  // ダイアログに表示するボタンと処理
			"ＯＫ": function() {
				var	url = 'sch.py?dns=' + $('#dns').val();
				var	tanto_code = $('#work_sch_edit > table > tbody > tr:eq(1) > td:eq(0) > div:eq(0)').text();
				url += '&TANTO_CODE=' + tanto_code;
				var	td = $('td',this);
				for(var i = 1; i < $(td).length; i++) {
					var	row = parseInt($('#work_sch_edit > table > caption').text()) - 1;
					var dst_tr = $('#work_sch > table > tbody > tr').eq(row);
					$('td', dst_tr).eq(i).html($(td).eq(i).html());
					var	ary = $(td).eq(i).attr('title').split(' ');
					url += '&CalDate' + i + '=' + ary[0];
					url += '&sTANTO_CODE' + i + '=' + ary[1];
					if(ary[1].trim() == '') {
						$(td).eq(i).attr('title', ary[0] + ' ' + tanto_code)
					}
					url += '&WorkDet' + i + '=';
					var	div = $('div:not(.off)', $(td).eq(i));
					for(var j = 0;j < $(div).length; j++) {
						if(j > 0) {
							url += ',';
						}
						url += $(div).eq(j).text();
					}
				}
//				$.toast({text: url,loader: false,hideAfter : 60000});
				fetch(url)
				.then((res) => {
					return res.json();
				})
				.then((json) => {
					var	flag = 'ok';
					for(var i = 1;i < 8;i++) {
						var	result = '';
						switch(i) {
						case 1:	result = json.result1; break;
						case 2:	result = json.result2; break;
						case 3:	result = json.result3; break;
						case 4:	result = json.result4; break;
						case 5:	result = json.result5; break;
						case 6:	result = json.result6; break;
						case 7:	result = json.result7; break;
						}
						if(result.indexOf('ok') == -1) {
							$.toast({text: result,loader: false,hideAfter : 60000});
							flag = 'error';
						}
					}
					if(flag == 'ok') {
						$.toast({text: '更新OK',loader: false,hideAfter : 1000});
					}
				})
				.catch((err) => {
					$.toast({text: err,loader: false,hideAfter : 60000});
				});
				$(this).dialog("close");
			},
            "キャンセル": function() {
				$(this).dialog("close");
            }
        },
        open: function(event, ui) {
//						$(this).dialog('option', 'title', '出荷日変更：' + $('#eODER_NO').val());
			$(this).dialog('option', 'title', );
			console.log('open' + this.id);
			$(this).dialog();
//			$('#navCnt').text('0');
        },
        close: function() {
//			$('#navCnt').text('0');
//			return false;
        }
    });
});
/*
欠品状況 stat2
*/
$(document).ready(function() {
	if($('#order_chg').val() > 0) {
		setConfig('#stat2_chg',2);
	} else {
		setConfig('#stat2_chg','');
	}
	setConfig('#stat2_scr','');
	setConfig('#stat2_fetch',600);
	$('#stat2').on('focus', function() {
		utimeOffset('#stat2_update','#' + this.id);
		$('#navChg').text($('#stat2_chg').val());
		$('#navScr').text($('#stat2_scr').val());
	});
	var timer = null;
	var	dt = [];
	$('#stat2_update').on('click', function() {
		console.log(this.id + '.click()');
		$(this).text($(this).text().replace('□','■'));
		// 同時処理防止
		if(fetchWait(this.id) == false) {
			var	url = 'stat2.py?dns=' + $('#dns').val();
			fetch(url)
			.then((res) => {
				// 同時処理解除
				fetchWait('');
				return res.json();
			})
			.then((json) => {
				var	tr = '';
				for ( var i = 0 ; i < json.data.length ; i++ ) {
					tr += '<tr>';
					tr += '<td rowspan="2"><span class="h6 text-top">' + json.data[i].SJCode + ' </span><span>' + json.data[i].Name + '</span></td>';
					tr += '<td>国内</td>';
					var	v;
					tr += '<td>' + ((v = json.data[i].cntPn1) != 0 ? v : '-') + '</td>';
					tr += '<td>' + ((v = json.data[i].cnt1) != 0 ? v : '-') + '</td>';
					tr += '<td>' + ((v = json.data[i].sumQty1) != 0 ? v : '-') + '</td>';
					tr += '</tr>';
					tr += '<tr>';
					tr += '<td>海外</td>';
					tr += '<td>' + ((v = json.data[i].cntPn2) != 0 ? v : '-') + '</td>';
					tr += '<td>' + ((v = json.data[i].cnt2) != 0 ? v : '-') + '</td>';
					tr += '<td>' + ((v = json.data[i].sumQty2) != 0 ? v : '-') + '</td>';
					tr += '</tr>';
				}
				if($(window).scrollTop() == 0) {
	//2018.12.28				$('#stat2 > table').floatThead('destroy');
				}
				$('#stat2 > table > tbody').find("tr").remove();
				$('#stat2 > table > tbody').append(tr);
				$(this).text('□更新.' + nowTM());
				utimeOffset('#' + this.id, '#stat2');
			})
			.catch((err) => {
				console.log(err);
				$.toast({text: '<div class="h7">' + url + '</div><div class="h7">fetch err:' + err + '</div>',loader: false});
			});
		}
		if($('#stat2_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					debug_toast('setTimeout:#stat2:' + $('#navId').text());
					$('#stat2_update').trigger("click");
				},$('#stat2_fetch').val() * 1000);
			}
		}
		return false;
	});
	if($('#stat2_chg').val() > 0) {
		$('#stat2_update').trigger("click");
		$('#stat2').removeClass('disable');
	} else {
		$('#stat2').addClass('disable');
	}
});
function dateYM(strDt) {
	if(strDt == '') {
		return strDt;
	}
	var strMMDD = strDt.slice(-4);
	return Number(strMMDD.slice(0,2)) + '/' + Number(strMMDD.slice(-2));
}
/*
出荷商品化待ち y_spot
*/
$(document).ready(function() {
	setConfig('#y_spot_chg','');
	setConfig('#y_spot_scr','');
	setConfig('#y_spot_fetch','');
	$('#y_spot').on('focus', function() {
		utimeOffset('#y_spot_update','#' + this.id);
		$('#navChg').text($('#y_spot_chg').val());
		$('#navScr').text($('#y_spot_scr').val());
	});
	var timer = null;
	$('#y_spot_update').on('click', function() {
		console.log(this.id + '.click()');
		// 同時処理防止
		if(fetchWait(this.id)) {
			return;
		}
		$(this).text($(this).text().replace('□','■'));
		var	url = 'y_spot.py?dns=' + $('#dns').val();
		url += '&Soko=' + $('#Soko').val();
		fetch(url)
		.then((res) => {
			// 同時処理解除
			fetchWait('');
			return res.json();
		})
		.then((json) => {
			var	tr = '';
			for ( var i = 0 ; i < json.data.length ; i++ ) {
				tr += '<tr>';
				tr += '<td class="rowNo">' + (i + 1) + '</td>';
//				<th>出荷日</th>
				tr += '<td class="Date">' + dateYM(json.data[i].SyukaDt) + '</td>';
//				<th>進捗</th>
				var	Stts = json.data[i].Stts;
				switch(Stts) {
				case '3':	Stts = '伝発待ち';	break;
				case '4':	Stts = '';	break;
				}
				if (json.data[i].Cnt > 1) {
					Stts = json.data[i].Cnt + '件';
				}
				tr += '<td class="Stts h4">' + Stts + '</td>';
//				<th>出荷先</th>
				tr += '<td class="Saki">';
				tr += json.data[i].aName;
				if (json.data[i].Cnt > 1) {
					tr += '<div>' + json.data[i].aName2 + '</div>';
				}
				tr += '</td>';
//				<th>倉庫</th>
//				tr += '<td class="SOKO_NAME"><div>' + json.data[i].ST_SOKO + '</div><div>' + json.data[i].SOKO_NAME + '</div></td>';
				tr += '<td';
				tr += ' class="SOKO_NAME soko-' + json.data[i].ST_SOKO + '"';
				tr += ' title="SOKO_NAME ' + json.data[i].ST_SOKO + '"';
				tr += '>';
//				tr += '<div class="' + json.data[i].ST_SOKO + '">';
				tr += json.data[i].SokoName;
//				tr += '</div>';
				tr += '</td>';
//				<th>品番</th>
				tr += '<td class="Pn">' + json.data[i].Pn;
				tr += '<div class="PName">' + json.data[i].HIN_NAME + '</div>';
				tr += '</td>'
//				過不足不足設定
				var	cls = 'number';
				var	cls0 = 'number';
				if(json.data[i].Qty > json.data[i].Qty0) {
					//不足
					cls += ' red';
					cls0 += ' red small';
				} else if(json.data[i].Qty < json.data[i].Qty0) {
					//不足なし
				} else {
					//イコール
					cls += ' yellow';
					cls0 += ' yellow';
				}
//				<th>出荷数</th>
				tr += '<td class="qty ' + cls + '">' + json.data[i].Qty + '</td>';
//				<th>不足数</th>
//				tr += '<td class="number">' + (json.data[i].Qty0 - json.data[i].Qty) + '</td>';
//				<th>在庫数 商済</th>
				tr += '<td class="qty0 ' + cls0 + '">' + (json.data[i].Qty0 == 0 ? '' : json.data[i].Qty0) + '</td>';
//				<th>在庫数 92</th>
				if(json.data[i].Qty <= (json.data[i].Qty0 + json.data[i].Qty92) && json.data[i].Qty92 > 0) {
					tr += '<td class="number qty92p">';
				} else {
					tr += '<td class="number qty92">'
				}
				tr += (json.data[i].Qty92 == 0 ? '' : json.data[i].Qty92);
				tr += '</td>';
//				<th>在庫数 未商</th>
				tr += '<td class="number qty1">' + (json.data[i].Qty1 == 0 ? '' : json.data[i].Qty1) + '</td>';
				tr += '</tr>';
			}
			tr += '<tr><td class="rowNo"></td><td colspan="9">　';
			if ( json.data.length >= 0) {
				tr += json.data.length + '件';
			} else {
				tr += '該当データがありません.';
			}
			tr += '</td></tr>';
			$('#y_spot > table > tbody').find("tr").remove();
			$('#y_spot > table > tbody').append(tr);
			$(this).text('□更新.' + nowTM());
			utimeOffset('#' + this.id, '#y_spot');
		})
		.catch((err) => {
			console.log(err);
			$.toast({text: '<div class="h7">' + url + '</div><div class="h7">fetch err:' + err + '</div>',loader: false});
		});
		if($('#y_spot_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					debug_toast('setTimeout:#y_spot:' + $('#navId').text());
					$('#y_spot_update').trigger("click");
				},$('#y_spot_fetch').val() * 1000);
			}
		}
		return false;
	});
	if($('#y_spot_chg').val() > 0) {
		$('#y_spot_update').trigger("click");
		$('#y_spot').removeClass('disable');
	} else {
		$('#y_spot').addClass('disable');
	}
});
/*
jcs_list
*/
var	jcs_list = [];
$(document).ready(function() {
	setConfig('#jcs_list_chg','');
	setConfig('#jcs_list_scr','');
	setConfig('#jcs_list_fetch','');
	$('#jcs_list').on('focus', function() {
		utimeOffset('#jcs_list_update','#' + this.id);
		$('#navChg').text($('#jcs_list_chg').val());
		$('#navScr').text($('#jcs_list_scr').val());
	});
	var timer = null;
	$('#jcs_list_update').on('click', function() {
		console.log(this.id + '.click()');
		// 同時処理防止
		if(fetchWait(this.id)) {
			return;
		}
		$(this).text($(this).text().replace('□','■'));
		var	url = 'jcs_list.py?dns=' + $('#dns').val();
		fetch(url)
		.then((res) => {
			// 同時処理解除
			fetchWait('');
			var contentType = res.headers.get("content-type");
			if(contentType && contentType.indexOf("application/json") !== -1) {
				return res.json();
			}
			$.toast({
				 header: url
				,text : 'Error content-type:' + contentType + '<br>' + res
				,loader: false
				,hideAfter : 60000
			});
		})
		.then((json) => {
			jcs_list = json.data;
			var	tr = '';
			var	tr_gray = '';
			for ( var i = 0 ; i < json.data.length ; i++ ) {
				if(i > 0) {
					if(json.data[i].DlvDt != json.data[i-1].DlvDt) {
						if(tr_gray == '') {
							tr_gray = 'gray';
						} else {
							tr_gray = '';
						}
					}
				}
				tr += '<tr class="' + tr_gray + '">';
				tr += '<td class="rowNo">' + (i + 1) + '</td>';
				//納入先
				tr += '<td class="DestCode ' + json.data[i].DestCode + '">';
				tr += json.data[i].DestCode;
				tr += '</td>';
				//場所
				tr += '<td class="Location">' + json.data[i].Location + '</td>';
				//納品書No
				tr += '<td class="NohinNo" title="' + json.data[i].SHIJI_NO + '">';
				tr += json.data[i].NohinNo + json.data[i].NohinNo2 + '</td>';
				//マツダ品番
				tr += '<td class="MazdaPn ' + json.data[i].CATEGORY_CODE + '"';
				tr += ' title="';
				tr += '\n' + json.data[i].N_CLASS_CODE;
				tr += '\n' + json.data[i].Pn;
				tr += '\n' + json.data[i].NameE;
				tr += '\n' + json.data[i].NameJ;
				tr += '\n' + json.data[i].SSpec;
				tr += '\n' + json.data[i].SType;
				tr += '\n' + json.data[i].GPn;
				tr += '\n' + json.data[i].PACKING_NO;
				tr += '\n' + json.data[i].CATEGORY_CODE;
				tr += '"';
				tr += '>';
				tr += json.data[i].MazdaPn + '</td>';
				//納入数
				tr += '<td class="Qty">' + json.data[i].Qty + '</td>';
				//在庫
				tr += '<td class="currQty">';
				if(json.data[i].currQty) {
					if(json.data[i].currQty > 0) {
						tr += json.data[i].currQty;
					}
				}
				tr += '</td>';
				//受注日
				tr += '<td class="Date" title="';
				tr += '\n' + json.data[i].HAKKO_DT;
				tr += '">';
				if (json.data[i].HAKKO_DT != '') {
					var	mm = json.data[i].HAKKO_DT.slice(4, 6);
					var	dd = json.data[i].HAKKO_DT.slice(6, 8);
					tr += parseInt(mm) + '/' + parseInt(dd);
				}
				tr += '</td>';
				//引取
				tr += '<td class="Hikitori" title="';
				tr += '\n' + json.data[i].Dt;
				tr += '\n' + json.data[i].iDestCode;
				tr += '\n' + json.data[i].DestName;
				tr += '\n' + json.data[i].IQty;
				tr += '">';
				if (json.data[i].Dt) {
					var	dt = json.data[i].Dt.replace(/-/g,'');
					if(dt >= json.data[i].HAKKO_DT) {
						var	mm = dt.slice(4, 6);
						var	dd = dt.slice(6, 8);
						tr += parseInt(mm) + '/' + parseInt(dd);
					}
				}
				tr += '</td>';
				//手配
				tr += '<td class="Tehai" title="';
				tr += '\n' + json.data[i].PRINT_DATETIME;
				tr += '\n' + json.data[i].TANTO_CODE;
				tr += '\n' + json.data[i].CANCEL_F;
				tr += '\n' + json.data[i].CANCEL_DATETIME;
				tr += '\n' + json.data[i].SHONIN_CODE;
				tr += '">';
				if(json.data[i].CANCEL_F == '1') {
					tr += 'ｷｬﾝｾﾙ';
				} else if(json.data[i].SHONIN_CODE != '') {
					tr += json.data[i].SHONIN_CODE;
				} else if(json.data[i].PRINT_DATETIME != '') {
					var	mm = json.data[i].PRINT_DATETIME.slice(4, 6);
					var	dd = json.data[i].PRINT_DATETIME.slice(6, 8);
					tr += parseInt(mm) + '/' + parseInt(dd);
				}
				tr += '</td>';
				//完了
				var	cls = "Kan"
				var	td = '';
				if(json.data[i].KAN_F) {
					if(json.data[i].KAN_F != '0') {
						var	mm = json.data[i].KAN_DT.slice(4, 6);
						var	dd = json.data[i].KAN_DT.slice(6, 8);
						td += parseInt(mm) + '/' + parseInt(dd);
//						tr += '○';
					} else if(json.data[i].HIN_CHECK_TANTO != '') {
//						tr += '&#10003';	//&#10004
						td += '<div class="box">作業中</div>';		//&#10004
//						cls += ' Working';
					}
				}
				tr += '<td class="' + cls + '" title="';
				tr += '\n' + json.data[i].KAN_F;
				tr += '\n' + json.data[i].KAN_DT;
				tr += '\n' + json.data[i].HIN_CHECK_TANTO;
				tr += '\n' + json.data[i].HIN_CHECK_DATETIME;
				tr += '\n' + json.data[i].HIN_CHECK_LABEL_CNT;
				tr += '\n' + json.data[i].HIN_CHECK_GENPIN_CNT;
				tr += '">';
				tr += td;
				tr += '</td>';
				//納品ﾁｪｯｸ
				tr += '<td class="NohinChk" title="' + json.data[i].EntID + '\n' + json.data[i].EntTm + '">';
				if(json.data[i].EntTm) {
					var	mm = json.data[i].EntTm.slice(5, 7);
					var	dd = json.data[i].EntTm.slice(8,10);
					tr += parseInt(mm) + '/' + parseInt(dd);
//					tr += '○';
				}
				tr += '</td>';
				//納入日
				tr += '<td class="DlvDt" title="' + json.data[i].DlvDt + '">';
				if(json.data[i].DlvDt.indexOf('-') > 0) {
					tr += json.data[i].DlvDt.slice(-5).replace('-','/');
				} else {
					var	mm = json.data[i].DlvDt.slice(4, 6);
					var	dd = json.data[i].DlvDt.slice(6, 8);
					tr += parseInt(mm) + '/' + parseInt(dd);
				}
				tr += '</td>';
				tr += '</tr>';
			}
			tr += '<tr><td class="rowNo"></td><td colspan="12">　';
			if ( json.data.length >= 0) {
				tr += json.data.length + '件';
			} else {
				tr += '該当データがありません.';
			}
			tr += '</td></tr>';
			$('#jcs_list > table > tbody').find("tr").remove();
			$('#jcs_list > table > tbody').append(tr);
			$(this).text('□更新.' + nowTM());
			utimeOffset('#' + this.id, '#jcs_list');
			$('#p_sshiji_cat_update').trigger("click");
		})
		.catch((err) => {
			console.log(err);
			$.toast({text: '<div class="h7">' + url + '</div><div class="h7">fetch err:' + err + '</div>',loader: false});
		});
		if($('#jcs_list_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					debug_toast('setTimeout:#jcs_list:' + $('#navId').text());
					$('#jcs_list_update').trigger("click");
				},$('#jcs_list_fetch').val() * 1000);
			}
		}
		return false;
	});
	if($('#jcs_list_chg').val() > 0) {
		$('#jcs_list_update').trigger("click");
		$('#jcs_list').removeClass('disable');
	} else {
		$('#jcs_list').addClass('disable');
	}
});
/*
jcs_summary
*/
$(document).ready(function() {
	setConfig('#jcs_summary_chg','');
	setConfig('#jcs_summary_scr','');
	setConfig('#jcs_summary_fetch','');
	$('#jcs_summary').on('focus', function() {
		$('#navChg').text($('#jcs_summary_chg').val());
		$('#navScr').text($('#jcs_summary_scr').val());
	});
	var timer = null;
	$('#jcs_summary .update').on('click', function() {
		$(this).text('◆' + $(this).text().slice(1));
		console.log($(this).attr("class") + '.click()');
		// 同時処理防止
		if(fetchWait('jcs_summary .update')) {
			return;
		}
		var	url = 'jcs_summary.py?dns=' + $('#dns').val();
		if($('#jcs_summary .noki').val()) {
			url += '&dt=' + $('#jcs_summary .noki').val();
		}
/*
		$.toast({
			text : url
			,loader: false
			,hideAfter : 3000
		});
*/
		console.log(url);
		fetch(url)
		.then((res) => {
			$(this).text('◇' + $(this).text().slice(1));
			console.log(res.headers.get("content-type"));
			// 同時処理解除
			fetchWait('');
			return res.json();
		})
		.then((json) => {
			var	stat1 = 0;
			var	stat2 = 0;
			var stat3 = 0;
			var stat4 = 0;
			var	tr = '';
			var	row = 0;
			for ( var i = 0 ; i < json.data.length ; i++ ) {
				var	ng = '';
				if(json.data[i].CANCEL_F == '1') {
					ng = 'ｷｬﾝｾﾙ';
				} else if(json.data[i].SHONIN_CODE != '') {
					ng = json.data[i].SHONIN_CODE;
				}
				if($('#jcs_summary .noki').val() == '' ) {
					if(ng != '') {
						continue;
					}
					$('#jcs_summary .noki').val(json.data[i].DlvDt);
				} else if($('#jcs_summary .noki').val() != json.data[i].DlvDt) {
					break;
				}
				tr += '<tr class="">';
				//#
				tr += '<td class="row">' + (++row) + '</td>';
				//納品書No
				tr += '<td class="no" title="';
				tr += '\n' + json.data[i].DlvDt;
				tr += '\n' + json.data[i].DestCode;
				tr += '\n' + json.data[i].SHIJI_NO;
				tr += '">';
				tr += json.data[i].NohinNo + json.data[i].NohinNo2 + '</td>';
				//マツダ品番
				tr += '<td class="pn" title="';
				tr += '\n' + json.data[i].N_CLASS_CODE;
				tr += '\n' + json.data[i].Pn;
				tr += '\n' + json.data[i].NameE;
				tr += '\n' + json.data[i].NameJ;
				tr += '\n' + json.data[i].SSpec;
				tr += '\n' + json.data[i].SType;
				tr += '\n' + json.data[i].GPn;
				tr += '">';
				tr += json.data[i].MazdaPn;
				tr += '</td>';
				tr += '<td class="pname">' + (json.data[i].NameJ != '' ? json.data[i].NameJ : json.data[i].NameE) + '</td>';
				//納入数
				tr += '<td class="qty">' + json.data[i].Qty + '</td>';
				//進捗
				tr += '<td class="stat" title="';
				tr += '\n' + json.data[i].CANCEL_F + ' ' + json.data[i].CANCEL_DATETIME;
				tr += '\n' + json.data[i].EntID + ' ' + json.data[i].EntTm;
				tr += '\n' + json.data[i].PRINT_DATETIME + ' ' + json.data[i].TANTO_CODE + ' ' + json.data[i].SHONIN_CODE;
				tr += '">';
				var	stat = 0;	//ｷｬﾝｾﾙ NG
				if(ng == '') {
					stat = 1;	//①手配前
					if(json.data[i].PRINT_DATETIME != '') {
						stat = 2;	//②商品化
						if(json.data[i].KAN_F != '0') {
							stat = 3;	//③納品チェック
							if(json.data[i].EntTm && json.data[i].EntTm != '') {
								stat = 4;	//④出荷レーン
							}
						}
					}
				}
				switch(stat) {
				case 1:	//①手配前
						stat1++;
						tr += '①';
						break;
				case 2:	//②商品化
						stat2++;
						tr += '②';
						break;
				case 3:	//③納品チェック
						stat3++;
						tr += '③';
						break;
				case 4:	//④出荷レーン
						stat4++;
						tr += '④';
						if(!json.data[i].EntID) {
							tr += '.';
						}
						break;
				case 0: //ｷｬﾝｾﾙ
						break;
				default:
						tr += stat;
						break;
				}
				tr += '</td>';
				tr += '<td>';
				switch(stat) {
				case 2:	//②商品化
						tr += json.data[i].HIN_CHECK_TANTO;
						break;
				case 0: //ｷｬﾝｾﾙ
						tr += ng;
						break;
				}
				tr += '</td>';
				tr += '</tr>';
			}
			$('#jcs_summary > table > tbody').find("tr").remove();
			$('#jcs_summary > table > tbody').append(tr);
			changeText($('#jcs_summary > .summary > .sumi > .qty'), stat2);
			changeText($('#jcs_summary > .summary > .zan  > .qty'), stat1);
			changeText($('#jcs_summary > .summary > .kan  > .qty'), stat4);
			changeText($('#jcs_summary > .summary2 > .stat1 > .qty'), stat1);
			changeText($('#jcs_summary > .summary2 > .stat2 > .qty'), stat2);
			changeText($('#jcs_summary > .summary2 > .stat3 > .qty'), stat3);
			changeText($('#jcs_summary > .summary2 > .stat4 > .qty'), stat4);
			$(this).text('□更新.' + nowTM());
		});
		if($('#jcs_summary_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					$('#jcs_summary .update').trigger("click");
				},$('#jcs_summary_fetch').val() * 1000);
			}
		}
		return false;
	});
	if($('#jcs_summary_chg').val() > 0) {
		$('#jcs_summary .update').trigger("click");
		$('#jcs_summary').removeClass('disable');
	} else {
		$('#jcs_summary').addClass('disable');
	}
	$('#jcs_summary > .summary > .sumi > .qty').on('click', function() {
		console.log('click:fadeInDown');
		$(this).fadeOut(500,function(){
//			$(this).text('100');
			$(this).fadeIn(1000,function(){
			});
		});
	});
});
/*
商品化予定
*/
$(document).ready(function() {
	setConfig('#package_plan_chg','');
	setConfig('#package_plan_scr','');
	setConfig('#package_plan_fetch','600');
	setConfig('#package_plan_file','');
	$('#package_plan').on('focus', function() {
		utimeOffset('#package_plan_update','#package_plan');
		$('#navChg').text($('#package_plan_chg').val());
		$('#navScr').text($('#package_plan_scr').val());
	});
	var timer = null;
	$('#package_plan_update').on('click', function() {
		console.log(this.id + '.click()');
		// 同時処理防止
		if(fetchWait(this.id)) {
			return;
		}
		$(this).text($(this).text().replace('□','■'));
		var	url = 'package_plan.py?dns=' + $('#dns').val();
		if($('#package_plan_file').val()) {
			url += '&filename=' + $('#package_plan_file').val();
		} else {
			if ($('#pref').text()) {
				url += '&filename=商品化予定_' + $('#pref').text() + '.xlsx';
			}
		}
		fetch(url)
		.then((res) => {
			// 同時処理解除
			fetchWait('');
			return res.json();
		})
		.then((json) => {
			var	tr = '';
			for ( var i = 0 ; i < json.data.length ; i++ ) {
				tr += '<tr>';
				tr += '<td class="rowNo">' + (i + 1) + '</td>';
				//予定日
				tr += '<td class="Date">';
				tr += json.data[i].Dt;
				tr += '</td>';
				if(json.data[i].Qty == '') {
					tr += '<td class="PnMemo" colspan="9">';
					tr += json.data[i].Pn;
					tr += '</td>';
					tr += '</tr>';
					continue;
				}
				//品番
				tr += '<td class="Pn">';
				tr += json.data[i].Pn;
				tr += '</td>';
				//品名
				tr += '<td class="PName">';
				tr += json.data[i].HIN_NAME;
				tr += '</td>';
				//数量
				var	strStat = json.data[i].stts;
/*
				if(json.data[i].zQty92 >= json.data[i].Qty) {
					strStat = "zQty92";
				} else if(json.data[i].sQty >= json.data[i].Qty) {
					strStat = "sQty";
				} else if(json.data[i].zQtySumi >= json.data[i].Qty) {
					strStat = "zQtyS";
				} else if(json.data[i].zQtyMi >= json.data[i].Qty) {
					strStat = "zQtyM";
				}
*/
				tr += '<td class="Qty ' + strStat + '">';
				tr += json.data[i].Qty;
				tr += '</td>';
				//見本
				tr += '<td class="Sample">';
				tr += json.data[i].Sample;
				tr += '</td>';
				//備考	状況
				tr += '<td class="Memo">';
				tr += json.data[i].Memo;
				tr += '</td>';
				//状況
				tr += '<td class="Stat ' + ((strStat == 'uQty' || strStat == 'sQty') ? strStat : '') + '">';
//				tr += '<td class="Stat ' + strStat + '">';
				if(strStat == 'uQty') {
					tr += '完了(' + json.data[i].uQty + ')';
				} else if(strStat == 'sQty') {
					tr += '完成品(' + json.data[i].sQty + ')';
				}
//				tr += json.data[i].QtyE;
				tr += '</td>';
				//在庫商済
				tr += '<td class="Qty ' + ((strStat == 'uQty' || strStat == 'sQty') ? strStat : '') + '">';
				tr += json.data[i].zQtySumi;
				tr += '</td>';
				//在庫92
				tr += '<td class="Qty ' + ((strStat == 'zQty92') ? strStat : '') + '">';
				tr += json.data[i].zQty92;
				tr += '</td>';
				//在庫未商
				tr += '<td class="Qty ' + ((strStat == 'zQtyM') ? strStat : '') + '">';
				tr += json.data[i].zQtyMi;
				tr += '</td>';
				tr += '</tr>';
			}
			tr += '<tr><td class="rowNo"></td><td colspan="10"><div class="info">';
			if ( json.data.length >= 0) {
				tr += json.data.length + '件';
			} else {
				tr += '該当データがありません.';
			}
			tr += '</div><div class="fileinfo">' + json.filename + ' ' + json.mtime.slice(0,16) + '</div>';
			tr += '</td></tr>';
			$('#package_plan > table > tbody').find("tr").remove();
			$('#package_plan > table > tbody').append(tr);
			$('th','#package_plan > table > thead > tr').eq(5).text(json.columns[3]);
			$('th','#package_plan > table > thead > tr').eq(6).text(json.columns[4]);
			$(this).text('□更新.' + nowTM());
			utimeOffset('#package_plan_update','#package_plan');
		})
		.catch((err) => {
			$(this).text('□更新.Error');
			console.log(err);
			$.toast({text: '<div class="h7"><div>.catch(err)</div><div>' + url + '</div><div>fetch err:' + err + '</div></div>',loader: false,hideAfter : 60000});
		});
		if($('#package_plan_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					$('#package_plan_update').trigger("click");
				},$('#package_plan_fetch').val() * 1000);
			}
		}
		return false;
	});
	if($('#package_plan_chg').val() > 0) {
		$('#package_plan_update').trigger("click");
		$('#package_plan').removeClass('disable');
	} else {
		$('#package_plan').addClass('disable');
	}
});
/*
入出荷予定
*/
$(document).ready(function() {
	setConfig('#inout_plan_chg','');
	setConfig('#inout_plan_scr','');
	setConfig('#inout_plan_fetch','600');
	$('#inout_plan').on('focus', function() {
		utimeOffset('#inout_plan_update','#inout_plan');
		$('#navChg').text($('#inout_plan_chg').val());
		$('#navScr').text($('#inout_plan_scr').val());
	});
	var timer = null;
	$('#inout_plan_update').on('click', function() {
		console.log(this.id + '.click()');
		// 同時処理防止
		if(fetchWait(this.id)) {
			return;
		}
		$(this).text($(this).text().replace('□','■'));
		var	url = 'inout_plan.py?dns=' + $('#dns').val();
		if ($('#pref').text()) {
			url += '&filename=入出荷予定_' + $('#pref').text() + '.xlsx';
		}
		fetch(url)
		.then((res) => {
			// 同時処理解除
			fetchWait('');
			return res.json();
		})
		.then((json) => {
			var	tr = '';
			for ( var i = 0 ; i < json.data.length ; i++ ) {
				tr += '<tr>';
				tr += '<td class="rowNo">' + (i + 1) + '</td>';
				for (var item in json.data[i]) {
					var	v = json.data[i][item] || '';
					tr += '<td class="c' + item + '">' + json.data[i][item] + '</td>';
				}
				tr += '</tr>';
			}
			$('#inout_plan > table > tbody').find("tr").remove();
			$('#inout_plan > table > tbody').append(tr);
			$(this).text('□更新.' + nowTM());
			utimeOffset('#inout_plan_update','#inout_plan');
		})
		.catch((err) => {
			$(this).text('□更新.Error');
			$('#inout_plan .info').text(err);
			console.log(err);
		});
		if($('#inout_plan_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					$('#inout_plan_update').trigger("click");
				},$('#inout_plan_fetch').val() * 1000);
			}
		}
		return false;
	});
	if($('#inout_plan_chg').val() > 0) {
		$('#inout_plan_update').trigger("click");
		$('#inout_plan').removeClass('disable');
	} else {
		$('#inout_plan').addClass('disable');
	}
});
/*
欠品入荷リスト
*/
$(document).ready(function() {
	setConfig('#short_plan_chg','');
	setConfig('#short_plan_scr','');
	setConfig('#short_plan_fetch','600');
	setConfig('#short_plan_file','');
	$('#short_plan').on('focus', function() {
		utimeOffset('#short_plan_update','#short_plan');
		$('#navChg').text($('#short_plan_chg').val());
		$('#navScr').text($('#short_plan_scr').val());
	});
	var timer = null;
	$('#short_plan_update').on('click', function() {
		console.log(this.id + '.click()');
		// 同時処理防止
		if(fetchWait(this.id)) {
			return;
		}
		$(this).text($(this).text().replace('□','■'));
		var	url = 'short_plan.py?dns=' + $('#dns').val();
		url += '&filename=' + $('#short_plan_file').val();
		fetch(url)
		.then((res) => {
			// 同時処理解除
			fetchWait('');
			return res.json();
		})
		.then((json) => {
			var	tr = '';
			for ( var i = 0 ; i < json.data.length ; i++ ) {
				tr += '<tr>';
				tr += '<td class="rowNo">' + (i + 1) + '</td>';
				for (var item in json.data[i]) {
					var	v = json.data[i][item] || '';
					tr += '<td class="c' + item + '">' + json.data[i][item] + '</td>';
				}
				tr += '</tr>';
			}
			$('#short_plan > table > tbody').find("tr").remove();
			$('#short_plan > table > tbody').append(tr);
			$('#short_plan .fileinfo').text($('#short_plan_file').val());
			$('#short_plan .info').text('');
			$(this).text('□更新.' + nowTM());
			utimeOffset('#short_plan_update','#short_plan');
		})
		.catch((err) => {
			$(this).text('□更新.Error');
			$('#short_plan .info').text((url) + ' ' + (err));
			console.log(err);
		});
		if($('#short_plan_fetch').val() > 0) {
			clearTimeout(timer);
			if($('#navId').text() != '#config_div') {
				timer = setTimeout(function() {
					$('#short_plan_update').trigger("click");
				},$('#short_plan_fetch').val() * 1000);
			}
		}
		return false;
	});
	if($('#short_plan_chg').val() > 0) {
		$('#short_plan_update').trigger("click");
		$('#short_plan').removeClass('disable');
	} else {
		$('#short_plan').addClass('disable');
	}
});
function changeText(q , v) {
	console.log('changeText(' + q + ',' + v + ')');

	if(q.text() != v) {
		q.fadeOut(500,function(){
			q.text(v);
			q.fadeIn(1000,function(){
			});
		});
//	    $(q).slideUp("slow", function() {
//			$(q).text(v);
//	    });
//		q.removeClass("fadeInDown");
//		q.text(v);
//		q.addClass("fadeInDown");
	}
}
//---------------------------------------------------------------
//音声再生
//---------------------------------------------------------------
function Speech(txt) {
//	sSpeech.Rate($('#spRate').text());
//	sSpeech.Pitch($('#spPitch').text());
//	sSpeech.Volume($('#spVolume').text());

//	sSpeech.say(txt);
//	sSpeech.saySample(txt,'Google 日本語');
//	$.toast({text:'sSpeech.voices.length=' + sSpeech.voices.length});
	sSpeech.saySample(txt.replace(' ','、'),$('#voiceName').val());
//	$.toast({text:$('#sp_text').val()});
//	$.toast({text:'Volume=' + sSpeech.Volume()
//				+ '<br>Rate=' + sSpeech.Rate()
//				+ '<br>Pitch=' + sSpeech.Pitch()
//			});
}
var sSpeech = {
    voices: null,
    synthes: null,
    init: function () {
        if (typeof (SpeechSynthesisUtterance) === "undefined") {
            return;
        }
        var o = this;
        o.synthes = new SpeechSynthesisUtterance();
        o.loadVoices();
    },
    loadVoices: function () {
        var o = this;
		if (!o.synthes) {
			return;
		}
		var	repeat  = setInterval(function() {
			if(o.synthes){
		        o.voices = window.speechSynthesis.getVoices();
	            // $voicesの中身を見てみる F12
//	            $.map(o.voices, function(n, i){console.log(n.name)});
				console.log('o.voices.length=' + o.voices.length);
				if(o.voices.length > 0) {
		            clearInterval(repeat);
					return;
				}
	        }
	    }, 300);
    },
    Volume: function (v) {
        /*
        ボリューム属性.volume
        この属性は、発話のためボリュームを指定します。0以上1以下の範囲で指定し、1が最大となります。
        */
        var o = this;
        if (!o.synthes) {
            return;
        }
        if (v) {
            if (v > 0) {
                v = v / 10;
            }
            o.synthes.volume = v;
        }
        return (o.synthes.volume);
    },
    Rate: function (v) {
        /*
        レート属性.rate
            この属性は、発話に対する発話速度を指定します。
            2は2倍の速さで、
            0.5は半分の速さです。
            0.1未満または10を超える値は厳しく禁止されています。
        */
        var o = this;
        if (!o.synthes) {
            return;
        }
        if (v) {
            o.synthes.rate = v;
        }
        return (o.synthes.rate);
    },
    Pitch: function (v) {
        /*
        ピッチ属性.pitch
            この属性は、発話に対するピッチを指定します。0が最低間隔であり、2が最高間隔となります。
        */
        var o = this;
        if (!o.synthes) {
            return;
        }
        if (v) {
            o.synthes.pitch = v;
        }
        return (o.synthes.pitch);
    },
    say: function (msgText) {
        /*
        テキスト属性.text
            この属性は、音声出力のために合成され、話されるテキストを指定します。
            テキストの最大長さは32,767文字に制限されます。
        lang属性.lang
            この属性は、有効なBCP 47言語タグを使用して、発話のための音声合成の言語を指定します。 
        */
        var o = this;
        if (!o.synthes) {
            return;
        }
//      o.sayStop();
        o.synthes.text = msgText;
        speechSynthesis.speak(o.synthes);
    },
    saySample: function (msgText, voiceName) {
        var o = this;
        if (!o.synthes) {
            return;
        }
        o.synthes.voice = null;
        if (voiceName) {
            for (var i = 0; i < o.voices.length; i++) {
                if (o.voices[i].name == voiceName)
                    o.synthes.voice = o.voices[i];
            }
        }
        o.sayStop();
        o.synthes.text = msgText;
        speechSynthesis.speak(o.synthes);
    },
    sayStop: function () {
        var o = this;
        if (!o.synthes) {
            return;
        }
        speechSynthesis.cancel();
    }
}
$(document).ready(function() {
	if(setConfig('#voiceName','Google 日本語') == '') {
		$('#voiceName').val('Google 日本語');
	}
	sSpeech.init();
	$('#sp_test').on('click', function() {
		Speech($('#sp_text').val());
	});
    // 音声一覧
	$("#loadVoice").on('click', function() {
	    $('#VoiceList').empty();
	    sSpeech.loadVoices();
	    var voices = sSpeech.voices;
	    if (!voices) {
	        $('#VoiceList').append("<li>undefined</li>");
	    } else {
	        if (voices.length == 0) {
	            $('#VoiceList').append("<li>なし</li>");
	        }
	        for (i = 0 ; i < voices.length ; i++) {
	            $('#VoiceList').append("<li>" + voices[i].lang + " " + voices[i].name + "</li>");
	        }
	    }
	    $('#VoiceList').trigger('create');
	});
	//朝一一回だけ
	var player = new Audio();
	$("#morning_play").on('click', function() {
		console.log(this.id + ':click');
//		var now = new Date();
//		var	dt = (now.getMonth() + 1) * 100 + now.getDate();
		var	dt = parseInt($('#morningDate').val()) % 10000;
		var	music = "sekaini";
		if (dt >= 101 && dt <= 131) {
			switch(dt % 6) {
			case 0:	music = "mbox_Pachelbel_canon";	break;
			case 1:	music = "Tchaikovsky_kurumi-march_mi";	break;
			case 2:	music = "Tchaikovsky_kurumi-rosi_mi";	break;
			case 3:	music = "Tchaikovsky_kurumi-konp_mi";	break;
			case 4:	music = "Tchaikovsky_kurumi-hana_mi";	break;
			case 5:	music = "Tchaikovsky_kurumi_hana_p";	break;
			}
		}
		if (dt >= 201 && dt <= 231) {
			switch(dt % 6) {
			case 0:	music = "CountryRoad";	break;
			case 1:	music = "Chiquitita";	break;
			case 2:	music = "HarryPotterPrologue";	break;
			case 3:	music = "E.T.";	break;
			case 4:	music = "OfficerAndAGentleman";	break;
			case 5:	music = "Singin'InTheRain";	break;
			}
		}
		if (dt >= 301 && dt <= 303) {
			music = "1103";	// ひなまつり
		}
		if (dt >= 304 && dt <= 331) {
			switch(dt % 4) {
			case 0:	music = "saboten";	break;
			case 1:	music = "haruyo.mid";	break;
			case 2:	music = "LostInLove";	break;
			case 3:	music = "AllOutOfLove";	break;
			}
		}
		if (dt >= 401 && dt <= 431) {
			switch(dt % 5) {
			case 0:	music = "HaveYouNeverBeenMellow";	break;
			case 1:	music = "LetMeBeThere";				break;
			case 2:	music = "SummerNights";				break;
			case 3:	music = "You'reTheOneThatIWant";	break;
			case 4:	music = "IHonestlyLoveYou";			break;
			}
		}
		if (dt >= 824 && dt <= 1031) {
			music = "majo_umi";	//
		}
		if (dt >= 1101 && dt <= 1130) {
			music = "The_Carpenters_-_Top_of_the_World";
			music = "The_Carpenters_-_Close_to_You";
		}
		if (dt >= 1201 && dt <= 1215) {
			if (dt % 2 === 0) {
				music = "Deniece_Williams_-_Let's_Hear_it_for_the_Boy";
			} else {
				music = "BJ_Thomas_-_Raindrops_Keep_Falling";
			}
		}
		if (dt >= 1223 && dt <= 1225) {
			music = "jinglehandbell";
		} else if (dt >= 1226) {
			switch(dt % 2) {
			case 0:	music = "mbox_Handel_hallelujah";	break;
			case 1:	music = "mbox_Beethoven_symp_9_gassyou";	break;
			}
		} else if (dt >= 1215 && dt <= 1231) {
			switch(dt % 3) {
			case 1:	music = "Sting";	break;
			case 2:	music = "Columbo";	break;
			case 0:	music = "SesameStreet";	break;
			}
		}
		//chime(music);
		player.src = 'sound/' + music + '.mp3';
		player.load();
		player.volume = 0.1;
		player.play();
		$.toast({heading : dt + ' ' + player.src + ' <input type="button" class="music-play" value="  ♪  "> '
				,text:'♪音楽を停止する場合は停止をクリック→<input type="button" class="music-stop" value="♪停止"> '
				,allowToastClose : true
				,loader: false
				,hideAfter : 3 * 60 * 1000
				,icon: 'info'
				});
		$(document).on("click", ".music-stop", function () {
			$.toast({text:'♪停止しました.'
						,loader: false
						,hideAfter : 30 * 1000
						});
//			player.pause();
			$('#stop').trigger('click');
			return false;
		});
		return false;
	});
	$(document).on("click", ".music-play", function () {
		player.play();
		$.toast({text:player.src
					,loader: false
					,hideAfter : 10 * 1000
					});
		return false;
	});
	$('#stop').on('click', function() {
		player.pause();
		return false;
	});
	$('#play_midi').on('click', function() {
		$.toast({text:'play:' + $('#midi').val()});
		MIDIjs.player_callback = function(ev) {
			$.toast({text:'player_callback:time=' + ev.time});
			if(ev.time > 3) {
				MIDIjs.player_callback = null;
			}
		};
		MIDIjs.play($('#midi').val());
	});
//	$('#stop2').on('click', function() {
	$("#good_morning").on('click', function() {
		console.log(this.id + ':click');
		if(sSpeech.voices.length > 0) {
			var now = new Date();
			var	nowText = (now.getMonth() + 1) + '月';
			if (now.getMonth() == 0 && now.getDate() <= 6) {
				nowText += now.getDate() + '日';
				nowText += '、' + [ "日", "月", "火", "水", "木", "金", "土" ][now.getDay()] + '曜日です。';
				Speech('新年あけましておめでとうございます。今年もよろしくお願い致します。');
			} else {
				nowText += now.getDate() + '日';
				nowText += '、' + [ "日", "月", "火", "水", "木", "金", "土" ][now.getDay()] + '曜日です。';
				Speech('おはようございます。' + nowText + '今日もいちにちよろしくお願いします。');
			}
		}
	});
	var today = dateYMD(new Date());
	if(setConfig('#morningDate','') < today) {
		$('#morningDate').val(today);
		$('#morningDate').change();
	    $('#morning_play').trigger('click');
		setTimeout(function() {
		    $('#good_morning').trigger('click');
		}, 15000);
//		setConfig('#morningDate', today);
//		$.toast({text:setConfig('#morningDate','') + ' ' + today});
	}
});
//メインループ：通知メッセージ
var intervalFunc = null;
$(document).ready(function() {
	$('#name').on('click', function(){
//		$('table').floatThead('destroy');
	});
	$('#tm').on('click', function(){
		$('#navWindow').toggle();
	});
	$('#navWindow').on('click', function(){
		$(this).toggle();
	});
	$('#play').on('click', function(){
		if($(this).text() == 'play_arrow') {
			$(this).text('pause');
			$.toast({text: '自動切替：OFF',	loader: false});
		} else {
			$(this).text('play_arrow');
			$.toast({text: '自動切替：ON',	loader: false});
		}
		return false;
	});
	$('#next,#prev').on('click', function(e){
//		$('nav').navChange(false);
		console.log('(#next,#prev).clise():' + this.id + ' e.which:' + e.which);
		if(e.which) {
			$('#navCnt').text('-1');
		}
		var	ary = $('a','nav');
		if(this.id == 'next') {
			ary = ary.get().reverse();
		}
		var	next = null;
		$(ary).each(function(){
			console.log('#next.click():' + $(this).attr('href') + ':' + location.hash + ':' + next);
			if($(this).attr('href') == location.hash) {
				if(next) {
					return false;	//break;
				}
			}
			if(!($($(this).attr('href')).hasClass('disable'))) {
				next = this;
			}
		});
		if(next) {
			$($(next).attr('href')).removeClass('slideInRight');
			$($(next).attr('href')).removeClass('slideInLeft');
			$($(next).attr('href')).addClass((this.id == 'next') ? 'slideInRight' : 'slideInLeft');
			$(next).trigger("click");
		}
		return false;
	});
/*
	$('#prev').on('click', function(){
		var	next = null;
		$('a','nav').each(function(){
			console.log('#next.click():' + $(this).attr('href') + ':' + location.hash + ':' + next);
			if($(this).attr('href') == location.hash) {
				if(next) {
					$(next).trigger("click");
					return false;
				}
			}
			if(!($($(this).attr('href')).hasClass('disable'))) {
				next = this;
			}
		});
		if(next) {
			$(next).trigger("click");
		}
/*
		var	prev = null;
		$('a','nav').each(function(i) {
			if($($(this).attr('href')).is(':visible')) {
				debug_toast($(this).text() + ':' + $(this).attr('href'));
				if(prev) {
					$(prev).trigger("click");
					return false;
				}
			}
			if(!($($(this).attr('href')).hasClass('disable'))) {
				prev = this;
			}
			if(i == $('a','nav').length - 1 ) {
				//最後
				if(prev) {
					$(prev).trigger("click");
					return false;
				}
			}
		});
		return false;
*/
//	});

	$('#auto').on('click', function(){
		if($(this).text() == 'stop') {
			$(this).text('auto');
		} else {
			$(this).text('stop');
		}
	});
	if($('#debug').val() > 0) {
		$('#navWindow').toggle();
	}
	$('nav > ul > li > a').on('click', function(){
//		debug_toast('click:' + $(this).attr('href') + ' ' + $(this).text());
//		$("html").stop();
//		$('#title').text($(this).text());
//		$('#navCnt').text('-1');
		var	target = $(this).attr('href');
		location.href = target;
		$('html').scrollTop(0);
		$('#navId').text(target);
		$(target).focus();
//		$(target).fadeIn(1000,function() {
//			$(target).trigger("focus");
//			$(target).focus();
//			$('html').scrollTop(0);
//			$('#navId').text(target);
//		});
		return false;
	});
//	$('a[href~="#package"]').trigger("click");
	setConfig('#notice','');
	setConfig('#notice_debug','');
	var	iNotice = setConfig('#IntervalNotice',15);
	$('#navChg').text(setConfig('#navChenge',15));
	$('#navScr').text(setConfig('#navScroll',''));
	if(location.hash) {
		$('.content').focus();
	} else {
		$('#next').trigger('click');
	}
//	$('#curr').trigger('click');
//	$('nav').navChange(true);
//	var	cnt = 0;
	if(iNotice > 0) {
		var	stopCnt = 0;
		intervalFunc = setInterval(function(){
//			$('#top').text($(window).scrollTop());

			if(location.hash == '#config_div') {
				clearInterval(intervalFunc);
			} else {
				if($($('#navId').text()).is(':focus') == false) {
					$($('#navId').text()).focus();
				}
			}
//			$('#tm').addClass('gif-load text-yellow');
			var now = new Date();
//			debug_toast(cnt);
//			cnt++;
			$('#tm').html(datetimeFormat(now) + '<span>.</span>');
			if($('#play').text() == 'pause') {
//				$('#tm').html(datetimeFormat(now) + '<span class="text-black">.</span>');
				stopCnt++;
			} else {
				stopCnt = 0;
				$('#navCnt').text(Number($('#navCnt').text()) + 1);
				if($('#navScr').text() != '') {
					if(Number($('#navCnt').text()) == Number($('#navScr').text())) {
						if($(window).scrollTop() == 0) {
							if($('#title').text() != '設定') {
								//画面スクロール
								$('#title').trigger("click");
							}
						}
					}
				}
				if(Number($('#navCnt').text()) >= Number($('#navChg').text())) {
					$('#navCnt').text(0);
//					$('nav').navChange(false);
					$('#next').trigger('click');
				}
				if($('#notice').val() != '') {
					url = 'notice.py?path=' + $('#notice').val();
					fetch(url, {cache: "no-store"})
					.then((res) => {
						console.log(url);
						console.log(res);
						if($('#notice_debug').val()) {
							$.toast({
									 heading : url
									,text : res.headers.get("content-type")
									,hideAfter : $('#notice_debug').val() * 1000
									});
						}
						return res.json();
					} )
					.then((json) => {
						console.log( url + ':json.title=' + json.title );
						if($('#notice_debug').val()) {
							$.toast({
									 heading : url
									,text : JSON.stringify(json)
									,hideAfter : $('#notice_debug').val() * 1000
									});
						}
						if (json.title.length > 0) {
							if (json.title.endsWith('.html')) {
								if(!($('#notice_debug').val())) {
									$.toast().reset('all');
								}
								$.toast({
									 heading : json.title
									,text : json.text
									,loader: false
									,bgColor : 'white'			// 背景色
									,textColor : 'black'			// 文字色
									,hideAfter : 60 * 1000
									,position : 'center-center'	// ページ内での表示位置
								});
							} else {
								$.toast({
									 text : '<div class="h1">' + json.title + '</div>' + '<pre class="h5">' + json.text + '</pre>'
									,loader: false
									,bgColor : 'white'			// 背景色
									,textColor : 'blue'			// 文字色
									,hideAfter : 60 * 1000
								});
							}
							chime(json.chime, json.volume);
							var	repeat = setInterval(function(){
								if(audio.paused) {
	//								Speech(json.title);
									Speech(json.speech);
									clearInterval(repeat);
								}
							},100);
						}
					})
					.catch((err) => {
						console.log(err);
						$.toast({text: '<div class="h7">' + url + '</div><div class="h7">fetch err:' + err + '</div>',loader: false, hideAfter : 60000});
					});
				}
			}
		},iNotice * 1000 /* 15000 */);
	}
	$(window).on('mousedown', function(e){
		console.log('mousedown:' + e.which);
//		$('#navChg').text(Number($('#navChg').text()) + 10);
//		$.toast().reset('all');
		$("html").stop();
/*
		$.toast({
			 text : '<div class="h7">mousedown:' + e.which + '</div>'
			,loader: false
			,bgColor : 'white'			// 背景色
			,textColor : 'blue'			// 文字色
//			,textAlign : 'center'		// テキストの配置
			,position : 'bottom-right'	// ページ内での表示位置
			,hideAfter : 1 * 1000
		});
*/
	});
	$(window).keydown(function(e){
//		$("html").stop();
		if($('#play').text() != 'pause') {
			$('#play').text('pause');
			$.toast({text: '自動切替：OFF',	loader: false});
		}
		if($('#navId').text() == '#config_div') {
			return;
		}
		switch(e.keyCode) {
		case 33:	//PageUp
		case 34:	//PageDown
		case 38:	//↑
		case 40:	//↓
			return;
			break;
		}
		debug_toast('<div class="h7">keydown:' + e.keyCode + '</div>');
		$.toast().reset('all');
		$('.centerWindow').hide();
/*
		$.toast({
			 text : '<div class="h7">keydown:' + e.keyCode + '(' + String.fromCharCode(e.keyCode) + ')</div>'
			,loader: false
			,bgColor : 'white'			// 背景色
			,textColor : 'blue'			// 文字色
//			,textAlign : 'center'		// テキストの配置
			,position : 'bottom-right'	// ページ内での表示位置
			,hideAfter : 10 * 1000
		});
*/
		switch(e.keyCode) {
		case 37:	//←
//		case 68:	//D ワイヤレスポインターの←
			$('#prev').trigger('click');
			return false;
			break;

/*
			$('#navCnt').text(-1);
			var	prev = null;
			$('a','nav').each(function(i) {
				if($($(this).attr('href')).is(':visible')) {
					debug_toast($(this).text() + ':' + $(this).attr('href'));
					if(prev) {
						$(prev).trigger("click");
						return false;
					}
				}
				if(!($($(this).attr('href')).hasClass('disable'))) {
					prev = this;
				}
				if(i == $('a','nav').length - 1 ) {
					//最後
					if(prev) {
						$(prev).trigger("click");
						return false;
					}
				}
			});
*/
		case 13:	//Enter
		case 39:	//→
		case 66:	//B ワイヤレスポインターの→
//			$('#navCnt').text(-1);
//			$('body').animate({
//				opacity: 0.25,    // 透明度0.25へ
//				left: '1000px'     // 現在位置から右へ50px移動
//				height: 'toggle'  // (カスタムアニメーション)
//			}, 5000, function() {
				// アニメーション完了後に実行する処理
//			$('nav').navChange(false);
			$('#next').trigger('click');
//			});
			break;
		case 38:	//↑
		case 40:	//↓
//			$('#navCnt').text(-1);
//			$('body').css('overflow-y','auto');
			break;
		}
	});
});
//スワイプ
$(document).ready(function() {
	var direction_x;
	var	position_x;
	var direction_y;
	var	position_y;
	$(window).on('touchstart', function(event){		//指が触れたか検知
		debug_toast(event.type + ':' + position_x + ' ' + direction_x + ':' + position_y + ' ' + direction_y);
		//スワイプ開始時の横方向の座標を格納
		position_x = event.originalEvent.touches[0].pageX;
	    direction_x = ''; //一度リセットする
		position_y = event.originalEvent.touches[0].pageY;
	    direction_y = ''; //一度リセットする
	});
	$(window).on('touchmove', function(event){		//指が動いたか検知
		debug_toast(event.type + ':' + position_x + ' ' + direction_x + ':' + position_y + ' ' + direction_y);
		//スワイプの方向（left／right）を取得
	    if (position_x - event.originalEvent.touches[0].pageX > 70) { // 70px以上移動しなければスワイプと判断しない
	      direction_x = 'left'; //左と検知
	    } else if (position_x - event.originalEvent.touches[0].pageX < -70){  // 70px以上移動しなければスワイプと判断しない
	      direction_x = 'right'; //右と検知
	    }
	    if (position_y - event.originalEvent.touches[0].pageY > 70) {
	      direction_y = 'down';
	    } else if (position_y - event.originalEvent.touches[0].pageY < -70){
	      direction_y = 'up';
	    }
	});
	$(window).on('touchend', function(event){		//指が離れたか検知
		debug_toast(event.type + ':' + position_x + ' ' + direction_x + ':' + position_y + ' ' + direction_y);
		var evt = $.Event('keydown');
		if (direction_x == 'right'){
			evt.keyCode = 37;
			$(window).trigger(evt);
		} else if (direction_x == 'left'){
			evt.keyCode = 39;
			$(window).trigger(evt);
		} else if (direction_y == 'up'){
			evt.keyCode = 38;
			$(window).trigger(evt);
		} else if (direction_y == 'down'){
			evt.keyCode = 40;
			$(window).trigger(evt);
		}
	});
	$(window).resize(function() {
		$('#screen').text(screen.width + ' x ' + screen.height);
		var	size = '(1024 x 576)(1280 x 720)(1440 x 810)(1920 x 1080)';
		$('#size').text(window.innerWidth + ' x ' + window.innerHeight + size);
	});
	$(window).trigger('resize');
	$('#navWindow').toggle();
	chime('silence');
});
