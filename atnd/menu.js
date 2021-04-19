/*
menu.js
2018.05.08
*/
$(document).ready(function() {
	$('nav>ul>li').click(function(e) {
		console.log('click():' + this.id);
		$('nav>ul>li.active').removeClass('active');
		$(this).addClass('active');
		$('#msg').html('');
	});
	$('#key').text(location.pathname.replace(/\/[^/]*$/, '') + '/');
	$("#menu").load('menu.html?q=' + getTs(),function(){
		$('#cname').text(localStorage.getItem($('#key').text() + '#cname'));
		console.log('load(menu.html).' + $('#cname').text());
		$('#dns').trigger('blur');
	});
	$("#dns").blur(function() {
		$('#dns_span').text($(this).val());
		var	key = $('#key').text() + '#cname';
		$('#cname').text(localStorage.getItem(key));
		var	req = 'jgyobu.py?dns=' + $(this).val() + '&jgyobu=0';
		console.log(req);
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
		console.log('getItem():' + key + ':' + val);
		if(!val) {
			switch(this.id) {
			case 'dns':	val = 'newsdc';	break;
			}
		}
		console.log('getItem():' + key + ':' + val);
		$(this).val(val);
	});
	$(document).on("change",'input[type="text"].config', function() {
		var	key = $('#key').text() + '#' + this.id;
		console.log('setItem():' + key + ':' + $(this).val());
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
$(document).ready(function() {
	// 初期値セット
	//設定：初期値
	//ユーザー名
	console.log('username...');
	$('#username').text('...');
	fetch('username.py').then((res) => {
		return res.json();
	}).then((json) => {
		console.log(json);
		$('#username').text(json.username);
	}).catch(function(err) {
		$('#username').text(err);
    });
});
