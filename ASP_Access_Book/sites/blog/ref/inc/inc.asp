﻿<SCRIPT RUNAT=SERVER LANGUAGE=JAVASCRIPT>
function getRequest(str,stype,requestType) {
	var requestStr
	if(requestType.toLowerCase() == 'form') {
		requestStr = new String(Request.form(str));
	} else if(requestType.toLowerCase() == 'querystring') {
		requestStr = new String(Request.querystring(str));
	} else if(requestType.toLowerCase() == 'servervariables') {
		requestStr = new String(Request.ServerVariables(str));
	} else {
		Response.write('错误的请求类型');
		Response.end();
		return null;
	}
	
	if(requestStr == 'undefined') {
		return '';
	}
	if(stype == 'int' && isNaN(requestStr)) {
			Response.write("参数类型错误")
			Response.end();
			return null;
	}
	if (requestStr.indexOf('\'') >= 0 || requestStr.indexOf(String.fromCharCode(0)) >= 0) {
		Response.write("提交内容中含有非法字符");
		Response.end();
	}
	return requestStr;
}


function  trim(str) {
	var tmpStr=new String(str);
    return tmpStr.replace(/(^\s*)|(\s*$)/g, '');
}
function hexdec(str)
{
    return parseInt(str,16);
}

function zeroFill(a,b)
{
    var z = hexdec(80000000);
    if (z & a)
    {
        a = a>>1;
        a &= ~z;
        a |= 0x40000000;
        a = a>>(b-1);
    }
    else
    {
        a = a >> b;
    }
    return (a);
}
function mix(a,b,c)
{
    a -= b; a -= c; a ^= (zeroFill(c,13));
    b -= c; b -= a; b ^= (a<<8);
    c -= a; c -= b; c ^= (zeroFill(b,13));
    a -= b; a -= c; a ^= (zeroFill(c,12));
    b -= c; b -= a; b ^= (a<<16);
    c -= a; c -= b; c ^= (zeroFill(b,5));
    a -= b; a -= c; a ^= (zeroFill(c,3));
    b -= c; b -= a; b ^= (a<<10);
    c -= a; c -= b; c ^= (zeroFill(b,15));
    var ret = new Array((a),(b),(c));
    return ret;
}

function GoogleCH(url,length)
{
    var init = 0xE6359A60;
    if (arguments.length == 1)
     length = url.length;   
    var a = 0x9E3779B9;
    var b = 0x9E3779B9;
    var c = 0xE6359A60;
    var k = 0;
    var len = length;
    var mixo = new Array();
    while(len >= 12)
    {
        a += (url[k+0] +(url[k+1]<<8) +(url[k+2]<<16) +(url[k+3]<<24));
        b += (url[k+4] +(url[k+5]<<8) +(url[k+6]<<16) +(url[k+7]<<24));
        c += (url[k+8] +(url[k+9]<<8) +(url[k+10]<<16)+(url[k+11]<<24));
        mixo = mix(a,b,c);
        a = mixo[0]; b = mixo[1]; c = mixo[2];
        k += 12;
        len -= 12;
    }
    c += length;
    switch(len)
    {
        case 11:
        c += url[k+10]<<24;
        case 10:
        c+=url[k+9]<<16;
        case 9 :
        c+=url[k+8]<<8;
        case 8 :
        b+=(url[k+7]<<24);
        case 7 :
        b+=(url[k+6]<<16);
        case 6 :
        b+=(url[k+5]<<8);
        case 5 :
        b+=(url[k+4]);
        case 4 :
        a+=(url[k+3]<<24);
        case 3 :
        a+=(url[k+2]<<16);
        case 2 :
        a+=(url[k+1]<<8);
        case 1 :
        a+=(url[k+0]);
    }
    mixo = mix(a,b,c);
    if (mixo[2] < 0)
    return (0x100000000 + mixo[2]);
    else
    return mixo[2];
}

function strord(s)
{
    var re = new Array();
    for(i=0;i<s.length;i++)
    {
        re[i] = s.charCodeAt(i);
    }
    return re;
}
function c32to8bit(arr32) 
{
    var arr8 = new Array(); 
    for(i=0;i<arr32.length;i++) 
    {
     for (bitOrder=i*4;bitOrder<=i*4+3;bitOrder++) 
     {
            arr8[bitOrder]=arr32[i]&255;
            arr32[i]=zeroFill(arr32[i], 8);
        }
    }
    return arr8;
}
function myfmod(x,y)
{
 var i = Math.floor(x/y);
        return (x - i*y);
}

function GoogleNewCh(ch)
{
 ch = (((ch/7) << 2) | ((myfmod(ch,13))&7));
 prbuf = new Array();
 prbuf[0] = ch;
 for(i = 1; i < 20; i++) {
       prbuf[i] = prbuf[i-1]-9;
 }
 ch = GoogleCH(c32to8bit(prbuf), 80);
 return ch;
  
}
function URLencode(sStr)
{
return encodeURIComponent(sStr).replace(/\+/g,"%2B").replace(/\//g,"%2F");
}
function getGoogleHostInfo(url){
    var reqgr = "info:" + url;
    var reqgre = "info:" + URLencode(url);
	//Response.Write(reqgr+"<br>"+reqgre);
    gch = GoogleCH(strord(reqgr));
    gch = "6" + GoogleNewCh(gch);
    var querystring = "http://toolbarqueries.google.com/search?client=navclient-auto&ch=" + gch + "&ie=UTF-8&oe=UTF-8&features=Rank:FVN&q=" + reqgre;
//Response.Write(querystring);
	var objXMLHTTP, xml;
	try{
		xml = Server.createObject("MSXML2.ServerXMLHTTP");
		xml.setTimeouts(5000,5000,5000,5000);
		 xml.open("GET", querystring, false);
		 xml.setRequestHeader( "User-Agent", "Mozilla/4.0 (compatible; GoogleToolbar 2.0.114-big; Windows XP 5.1)" );
		 xml.send();
		 //Response.Write(xml.responseText);
		 //Response.Write(xml.responseBody);
		 return xml.responseText;
	}catch(exception){
		return("Rank_0:0:0");
	}
}
function getPageRank(temp){
  var foo = temp.match(/Rank_.*?:.*?:(\d+)/i);
  var pr = (foo) ? foo[1] : '0';
  return pr;
}
function getDirectory(temp){
  var  foo = temp.match(/FVN_.*?:.*?:(?:Top\/)?([^\s]+)/i);
  var cat = (foo) ? foo[1] : "";
  if(cat!="")cat="http://directory.google.com/Top/"+cat
  return cat;
}

</SCRIPT>