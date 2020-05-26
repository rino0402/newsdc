/*
menu.js
2018.05.08
*/
$(document).ready(function() {
	$('nav>ul>li').click(function(e) {
		console.log('click():' + this.id);
		$('nav>ul>li.active').removeClass('active');
		$(this).addClass('active');
	});
	$('#key').text(location.pathname.replace(/\/[^/]*$/, '') + '/');
	$("#menu").load('menu.html?q=' + getTs(),function(){
		$('#cname').text(localStorage.getItem($('#key').text() + '#cname'));
		console.log('load(menu.html).' + $('#cname').text());
		$('#dns').trigger('blur');
	});
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
			case 'limit':	val = 1000;	break;
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
});
function getTs() {
    var now = new Date();
    var res = "" + now.getFullYear() + padZero(now.getMonth() + 1) + padZero(now.getDate()) + padZero(now.getHours()) + 
        padZero(now.getMinutes()) + padZero(now.getSeconds());
    return res;
}
//先頭ゼロ付加
function padZero(num) {
    return (num < 10 ? "0" : "") + num;
}
function dns() {
	var	dns = 'newsdc';
	if(dns = location.search.substring(1)) {
	}
	return dns;
}
$(document).ready(function() {
	$("#config").load('config.html', function(response, status, xhr) {
		console.log('load:config.html');
		console.log('response:' + response);
		console.log('status:' + status);
		console.log('xhr:' + xhr);
	});
	// 初期値セット
	//設定：初期値
	$('#pref').text(location.search.substring(1));
//	var dns = storage('#dns');
	var dns = location.search.substring(1);
	console.log('dns:' + dns);
	if(dns == '') {
		dns = location.pathname;
		dns = dns.split('/')[1];
//		if(location.host == 'w0') {
//			dns = 'newsdc4';
//		}
	}
	console.log('dns:' + dns);
	$('#dns').text(dns);
	console.log('#dns:' + $('#dns').text());
	//設定：変更メソッド
//	$('.config').change(function() {
	$(document).on("change",".config", function() {
		console.log( 'change() ' + this.id + ':' + $(this).val());
		var id = '#' + this.id;
		storage(id,$(this).val());
	});
});
//localStorage 保存
function storage(n,v) {
	console.log( 'storage() ' + n + ':' + v);
	if(typeof v === 'undefined') {
		v = null;
	}
//	n = location.pathname + '.' + n;
	var	stg = localStorage;
	if(v == null) {
		v = stg.getItem(n,v);
	} else {
		stg.setItem(n,v);
	}
	console.log( 'storage().' + n + ':' + v);
	return v;
}
