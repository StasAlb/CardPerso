function ShowError(errorText, onOK) {
    HideWait();
    $("#popup_error_text").html(errorText);
    $("#popup_error_text").select();
    $("#popup_error").dialog({
        resizable: false,
        modal: true,
        close: function (event, ui) { $(this).dialog("destroy"); return true; },
        buttons: {
            OK: function () {
                $(this).dialog("close");
                if (typeof (onOK) == 'function') onOK();
            }
        }
    });
    $("#popup_error").keydown(function (event) {
        if (event.keyCode == 13) {
            $(this).parent().find("button:eq(0)").trigger("click");
            return false;
        }
    });
}

function ShowMsg(title, msgText, onOK) {
    HideWait();
    $("#popup_msg_text").html(msgText);
    $("#popup_msg").dialog({
        title: title,
        resizable: false,
        modal: true,
        width: 450,
        close: function (event, ui) { $(this).dialog("destroy"); return true; },
        buttons: {
            OK: function () {
                $(this).dialog("close");
                if (typeof (onOK) == 'function') onOK();
            }
        }
    });
    $("#popup_msg").keydown(function (event) {
        if (event.keyCode == 13) {
            $(this).parent().find("button:eq(0)").trigger("click");
            return false;
        }
    });

}

function ShowMessage(msg, onOK, onCancel) {
    ShowMsg("Информация", msg, onOK,onCancel);
}

function ShowQuestion(title, msgText, onOK, onCancel) {
    HideWait();
    $("#popup_question_text").html(msgText);
    $("#popup_question").dialog({
        title: title,
        resizable: false,
        modal: true,
        width: 400,
        close: function (event, ui) { $(this).dialog("destroy"); return true; },
        buttons: {
            Да: function () {
                $(this).dialog("close");
                if (typeof (onOK) == 'function') onOK();
            },
            Нет: function () {
                $(this).dialog("close");
                if (typeof (onCancel) == 'function') onCancel();
            }
        }
    });

    $("#popup_question").keydown(function (event) {
        if (event.keyCode == 13) {
            $(this).parent().find("button:eq(1)").trigger("click");
            return false;
        }
    });
    //setTimeout('$("#popup_question").parent().find("button:eq(1)").focus();',300);
    $("#popup_question").parent().find("button:eq(1)").focus();
}



var isWaitOpen = false;
var timeOutWait = null;

function ShowWait(msgText, params) {
    var title = "Ожидание завершения операции",
		width = 400,
		height = 100,
		timeOut = 0;

    if (typeof (params) != "undefined" && params != null) {
        if (typeof (params.title) != "undefined" && params.title != null) title = params.title;
        if (typeof (params.width) != "undefined" && params.width != null) width = params.width;
        if (typeof (params.height) != "undefined" && params.height != null) height = params.height;
        if (typeof (params.timeOut) != "undefined" && params.timeOut != null) timeOut = params.timeOut;
    }
    var m = (typeof (msgText) != 'undefined' && msgText != null) ? msgText : "Подождите пожалуйста...";
    $("#popup_wait_text").text(m);
    if (isWaitOpen) {
        if (timeOutWait != null) {
            clearTimeout(timeOutWait);
            timeOutWait = null;
        }
        return;
    }
    var popup_wait = $("#popup_wait");
    popup_wait.dialog({
        title: title, autoOpen: (timeOut < 1) ? true : false, resizable: false, modal: true,
        closeOnEscape: false, width: width, height: height
    }).parent().find(".ui-dialog-titlebar-close").hide();
    if (timeOut < 1) {
        //popup_wait.dialog("open");
        isWaitOpen = true;
    }
    else timeOutWait = setTimeout("$('#popup_wait').dialog('open');isWaitOpen=true;", timeOut);
}

function HideWait() {
   if (timeOutWait != null) {
        clearTimeout(timeOutWait);
        timeOutWait = null;
    }
    var popup_wait = $("#popup_wait");
    if (isWaitOpen) { popup_wait.dialog("close"); popup_wait.dialog("destroy"); isWaitOpen = false; }
}

function ShowWaitText(msgText) {
    if (typeof (msgText) != 'undefined' && msgText != null) $("#popup_wait_text").text(msgText);
}



function ajaxJson(url, jsonReq, onOK, onError, isAsync)
{
    var async = true;
    if (typeof (isAsync) != 'undefined') async = false;
    $.ajax({
        "async": async,
        "type": "POST",
        "url": url,
        "data": jsonReq,
        "success": function ok(json)
        {
            if (json!=null && typeof(json)!='undefined' && typeof(json.SESSION_END)!='undefined' && json.SESSION_END == true)
            {
                if (typeof (SESSION_END) == 'function') SESSION_END();
                return;
            }
            if (typeof (onOK) == 'function') onOK(json);
        },
        "dataType": "json",
        "cache": false,
        "error": function (xhr, error, thrown)
        {
            if (typeof (onError) == 'function') onError(error);
        }
    });
}

function getLocalCookie(name) 
{ 
   var matches = document.cookie.match(new RegExp("(?:^|; )" + name.replace(/([\.$?*|{}\(\)\[\]\\\/\+^])/g, '\\$1') + "=([^;]*)")); 
   return matches ? decodeURIComponent(matches[1]) : null; 
} 

function setLocalCookie(name,value,options) 
{ 
   options = options || {}; 
   var expires = options.expires; 
   if(typeof(expires)=="number" && expires) 
   { 
        var d = new Date(); 
        d.setTime(d.getTime() + expires*1000); 
        expires = options.expires = d; 
   } 
   if (expires && expires.toUTCString) 
   {  
       options.expires = expires.toUTCString(); 
   } 
   value = encodeURIComponent(value); 
   var updatedCookie = name + "=" + value; 
   for(var propName in options) 
   { 
       updatedCookie += "; " + propName; 
       var propValue = options[propName];     
       if (propValue !== true) 
       {  
           updatedCookie += "=" + propValue; 
       } 
   } 
   document.cookie = updatedCookie; 
} 


function sha1(str) {
  //  discuss at: http://phpjs.org/functions/sha1/
  // original by: Webtoolkit.info (http://www.webtoolkit.info/)
  // improved by: Michael White (http://getsprink.com)
  // improved by: Kevin van Zonneveld (http://kevin.vanzonneveld.net)
  //    input by: Brett Zamir (http://brett-zamir.me)
  //   example 1: sha1('Kevin van Zonneveld');
  //   returns 1: '54916d2e62f65b3afa6e192e6a601cdbe5cb5897'

  var rotate_left = function(n, s) {
    var t4 = (n << s) | (n >>> (32 - s));
    return t4;
  };

  /*var lsb_hex = function (val) {
   // Not in use; needed?
    var str="";
    var i;
    var vh;
    var vl;

    for ( i=0; i<=6; i+=2 ) {
      vh = (val>>>(i*4+4))&0x0f;
      vl = (val>>>(i*4))&0x0f;
      str += vh.toString(16) + vl.toString(16);
    }
    return str;
  };*/

  var cvt_hex = function(val) {
    var str = '';
    var i;
    var v;

    for (i = 7; i >= 0; i--) {
      v = (val >>> (i * 4)) & 0x0f;
      str += v.toString(16);
    }
    return str;
  };

  var blockstart;
  var i, j;
  var W = new Array(80);
  var H0 = 0x67452301;
  var H1 = 0xEFCDAB89;
  var H2 = 0x98BADCFE;
  var H3 = 0x10325476;
  var H4 = 0xC3D2E1F0;
  var A, B, C, D, E;
  var temp;

  // utf8_encode
  str = unescape(encodeURIComponent(str));
  var str_len = str.length;

  var word_array = [];
  for (i = 0; i < str_len - 3; i += 4) {
    j = str.charCodeAt(i) << 24 | str.charCodeAt(i + 1) << 16 | str.charCodeAt(i + 2) << 8 | str.charCodeAt(i + 3);
    word_array.push(j);
  }

  switch (str_len % 4) {
  case 0:
    i = 0x080000000;
    break;
  case 1:
    i = str.charCodeAt(str_len - 1) << 24 | 0x0800000;
    break;
  case 2:
    i = str.charCodeAt(str_len - 2) << 24 | str.charCodeAt(str_len - 1) << 16 | 0x08000;
    break;
  case 3:
    i = str.charCodeAt(str_len - 3) << 24 | str.charCodeAt(str_len - 2) << 16 | str.charCodeAt(str_len - 1) <<
      8 | 0x80;
    break;
  }

  word_array.push(i);

  while ((word_array.length % 16) != 14) {
    word_array.push(0);
  }

  word_array.push(str_len >>> 29);
  word_array.push((str_len << 3) & 0x0ffffffff);

  for (blockstart = 0; blockstart < word_array.length; blockstart += 16) {
    for (i = 0; i < 16; i++) {
      W[i] = word_array[blockstart + i];
    }
    for (i = 16; i <= 79; i++) {
      W[i] = rotate_left(W[i - 3] ^ W[i - 8] ^ W[i - 14] ^ W[i - 16], 1);
    }

    A = H0;
    B = H1;
    C = H2;
    D = H3;
    E = H4;

    for (i = 0; i <= 19; i++) {
      temp = (rotate_left(A, 5) + ((B & C) | (~B & D)) + E + W[i] + 0x5A827999) & 0x0ffffffff;
      E = D;
      D = C;
      C = rotate_left(B, 30);
      B = A;
      A = temp;
    }

    for (i = 20; i <= 39; i++) {
      temp = (rotate_left(A, 5) + (B ^ C ^ D) + E + W[i] + 0x6ED9EBA1) & 0x0ffffffff;
      E = D;
      D = C;
      C = rotate_left(B, 30);
      B = A;
      A = temp;
    }

    for (i = 40; i <= 59; i++) {
      temp = (rotate_left(A, 5) + ((B & C) | (B & D) | (C & D)) + E + W[i] + 0x8F1BBCDC) & 0x0ffffffff;
      E = D;
      D = C;
      C = rotate_left(B, 30);
      B = A;
      A = temp;
    }

    for (i = 60; i <= 79; i++) {
      temp = (rotate_left(A, 5) + (B ^ C ^ D) + E + W[i] + 0xCA62C1D6) & 0x0ffffffff;
      E = D;
      D = C;
      C = rotate_left(B, 30);
      B = A;
      A = temp;
    }

    H0 = (H0 + A) & 0x0ffffffff;
    H1 = (H1 + B) & 0x0ffffffff;
    H2 = (H2 + C) & 0x0ffffffff;
    H3 = (H3 + D) & 0x0ffffffff;
    H4 = (H4 + E) & 0x0ffffffff;
  }

  temp = cvt_hex(H0) + cvt_hex(H1) + cvt_hex(H2) + cvt_hex(H3) + cvt_hex(H4);
  return temp.toLowerCase();
}



  


