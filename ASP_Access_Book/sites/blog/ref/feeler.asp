<!--#include file="syscore.asp"-->
<script RUNAT=SERVER LANGUAGE=JAVASCRIPT>
/**********************************************************
* IYI web Stat. system v1.12 - Copyright (c) 2005 www.iyi.cn.
* This copyright notice MUST stay intact for use (see license.txt).
* 使用时必须保证本版本信息源有格式.
* For full source code and docs, visit http://blog.iyi.cn/user/david/
* 下载完整源码或者版本升级请访问http://blog.iyi.cn/user/david/
**********************************************************
* File:			observer.asp
* Version:		1.12
* Date:			2005-7-4
* Written by	David
**********************************************************
* blog:		David: http://blog.iyi.cn/david
* Email:	davidnick@126.com
* QQ:		76522970
***********************************************************/

//public variable obUrl
var obUrl;		//remote url
obUrl = escape(getRequest('HTTP_REFERER','str','servervariables'));

//class Feeler()
function Feeler() {
	
	//private variables
	var showType;
	var feelerStr;
	showType = getRequest('showType','int','querystring');

	//javascript on client
	feelerStr = "" +
	"var obReferer = document.referrer;\n" +
	"var isPublic = " + core.isPublic + ";\n" +
	"var serverUrl = \"" + core.serverUrl + "\";\n" +
	"var obHost = escape(document.location.protocol + \"//\" + document.location.host);\n" +
	"if ((obReferer == \"\") || (obReferer.indexOf(document.location.host)>0)) {\n" +
	"	obReferer = obHost;\n" +
	"}else{\n" +
	"	obReferer = escape(obReferer);\n" +
	"}\n" +
	"function fResponse()\n" +
	"{\n" +
	"	var strURL = serverUrl + \"observer.asp?obShowType=" + showType + "&obUrl=" + obUrl + "&obReferer=\" + obReferer + \"&obHost=\" + obHost + \"\";\n" +
	//"	document.write(\".\");" +
	//"	divReferer.innerHTML = \"Loading...\";\n" +
	//"	document.write(strURL);\n" +
	"	document.write(\"<sc" + "ript src=\" + strURL + \"></sc" + "ript>\");\n" +
	"}\n" +
	"//--------------------------------------------------------\n" +
	"if ((isPublic) || (serverUrl.indexOf(document.location.host)>0)) {\n" +
	"	fResponse();\n" +
	"}\n";

	//method sendFeeler
	this.sendFeeler = function() {
		Response.write(feelerStr);
	}
}

//main function
if(obUrl != '') {
	var myFeeler = new Feeler();
	myFeeler.sendFeeler();
}else{
	//Response.write('<DIV id="divReferer"></DIV>');
	Response.write('<s' + 'cript type="text/javascript" src="feeler.asp?showType=3"></s' + 'cript>');
}
</script>