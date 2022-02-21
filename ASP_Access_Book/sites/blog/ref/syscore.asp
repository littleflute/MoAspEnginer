<%@LANGUAGE=JAVASCRIPT codepage=65001 EnableSessionState=False%>
<!--#include file="inc/inc.asp"-->
<SCRIPT RUNAT=SERVER LANGUAGE=JAVASCRIPT>
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
function SysCore() {
	this.version = "1.11";
	this.isPublic = 1;
	this.authorInfo = "<a href = \"http://blog.iyi.cn/david\">David</a>";
	this.introPage = "";
	this.brandName = "IYI web Stat. system.";
	this.serverName = getRequest("SERVER_NAME","str","servervariables");	//Domain widthout port
	this.serverHost = getRequest("HTTP_HOST","str","servervariables");		//Domain width port
	this.serverUri = getRequest("URL","str","servervariables");
	var pos = this.serverUri.indexOf("feeler.asp");
	this.serverUri = this.serverUri.substr(0,pos);
	this.serverUrl = "http://" + this.serverHost + this.serverUri;
	this.initialize = function() {						//暂时保留，没有使用。
		this.engine;
		this.engine = ""; // 根据必要的信息创建字符串。
		this.engine += ScriptEngine() + " Version ";
		this.engine += ScriptEngineMajorVersion() + ".";
		this.engine += ScriptEngineMinorVersion() + ".";
		this.engine += ScriptEngineBuildVersion();
		Application("ENGINE") = this.engine;
		Application("VERSION") = "1.11";
		Application("AUTHER_INFO") = "<a href = \"http://blog.iyi.cn/david\">David</a>";
		Application("INTRO_PAGE") = "http://blog.iyi.cn/user/david/archives/2005/01/189.html";
		Application("BRAND_NAME") = "IYI web Stat. system.";
		Application("SERVER_NAME") = getRequest("SERVER_NAME","str","servervariables");	//Domain widthout port
		Application("SERVER_HOST") = getRequest("HTTP_HOST","str","servervariables");	//Domain width port
	}
}
var core = new SysCore();

</SCRIPT>