<!--#include file="syscore.asp"-->
<!--#include file="inc/conn.asp"-->
<SCRIPT RUNAT=SERVER LANGUAGE=JAVASCRIPT>
Response.Charset = 'utf-8';
Response.buffer = true;
function Observer() {
	var obSiteID;
	var obPageID;
	var obUrl;
	var obHost;
	var obReferer;
	var obCount;
	var obStartTime;
	var obShowType;
	var obPRValue;
	var obPRTime;
	var hasSite;
	var hasPage;
	var showStr;
	var sql;
	var rs;
	var conn;
	var nowTime;
	//var obTitle;
	this.initialize = function() {
		nowTime = (new Date()).toLocaleString()
		var re = new RegExp("^https{0,1}://[^<|>|'| ]+","gi");
		//obShowType = 3;
		obShowType =  getRequest('obShowType','int','querystring');
		//obUrl = 'http://localhost1';
		obUrl = new String(re.exec(getRequest('obUrl','str','querystring')));
		re.compile("^https{0,1}://[^<|>|'| ]+","gi");
		//obReferer = 'http://localhost1';
		obReferer = new String(escape(re.exec(getRequest('obReferer','str','querystring'))));
		re.compile("^https{0,1}://[^<|>|'| ]+","gi");
		//obHost = 'localhost1';
		obHost = new String(re.exec(getRequest('obHost','str','querystring')));
		obStartTime = new Date().getTime();
		showStr = '';
		
		conn = new Conn();
		conn.getConn();
		rs = Server.createObject('ADODB.RecordSet');
		sql = 'select * from localsite where siteDomain = \'' + obHost + '\'';
		rs.open(sql,conn.conn,1,3);
		hasSite = rs.recordcount;
		if(!hasSite) {
			rs.addNew;
			rs('sitedomain') = obHost;
			rs('sitename') = obHost;
			rs('sitecount') = 1;
			rs('sitetime') = nowTime;
		}else{
			rs('sitecount') += 1;
			rs('sitetime') = nowTime;
			rs.update;
		}
		rs.update;
		obSiteID = new Number(rs(0));
		rs.close();

		sql = 'select * from localpage where sitePage = \'' + obUrl + '\' and siteID = ' + obSiteID + '';
		rs.open(sql,conn.conn,1,3);
		hasPage = rs.recordcount;
		var nowDate = Math.floor(obStartTime/86400000);
		if(!hasPage){
			obPRValue = this.getPR();
			rs.addNew;
			rs('siteID') = obSiteID;
			rs('sitePage') = obUrl;
			rs('pageCount') = 1;
			rs('pageTime') = nowTime;
			rs('pagerank') = obPRValue;
			rs('prTime') = nowDate;
		}else{
			var prVaule = new String(rs('pagerank'));
			var prTime = new Number(rs('prTime'));
			if(nowDate - prTime >= 1){
				obPRValue = this.getPR();
				obPRTime = nowDate;
			}else{
				obPRValue = prVaule;
				obPRTime = prTime;
			}
			rs('pageCount') += 1;
			rs('pageTime') = nowTime;
			rs('pagerank') = obPRValue;
			rs('prTime') = obPRTime;
		}
		rs.update;
		obCount = new Number(rs(3));
		obPageID = new Number(rs(0));
		rs.close();

		sql = 'select * from referrer where pageID = ' + obPageID + ' and refUrl = \'' + obReferer + '\'';
		rs.open(sql,conn.conn,1,3);
		if(rs.recordcount) {
			rs('refCount') += 1;
			rs('refTime') = nowTime;
		}else{
			rs.addNew;
			rs('pageID') = obPageID;
			rs('refUrl') = obReferer;
			rs('refDomain') = "";
			rs('refName') = "";
			rs('refCount') = 1;
			rs('refTime') = nowTime;
		}
		rs.update;
		rs.close();

	}
	
	this.getObName = function() {
		return obName;
	}
	
	this.getObHost = function() {
		return obHost;
	}
	
	this.getObUrl = function() {
		return obUrl;
	}
	
	this.getOSTime = function() {
		return obStartTime;
	}
	
	this.showInfo = function() {
		//inner private function showPR
		var showPR = function() {
			showStr += 'd' + 'ocument.write(\'<link rel="stylesheet" href="http://stat.iyi.cn/inc/pr.css" type="text/css"/><div id="prouter"><div id="prleft"><div id="linner">PageRank</div><div id="pr"><div id="prinner" style="width:' + obPRValue + '0%"></div></div></div><div id="prright"><div id="rinner">' + obPRValue + '</div></div></div>\');\n';
		}
		
		//inner privaete function showCount
		var showCount = function() {
			showStr += 'd' + 'ocument.write(\'<div id="prouter">This page has bean visited ' + obCount + ' times\');\n';
			showStr += 'd' + 'ocument.write(\'<br/>\');\n';
			if(hasPage) {
				sql = 'select top 10 * from referrer where pageID = ' + obPageID + ' order by refTime desc';
				rs.open(sql,conn.conn,1,3);
				var c = 0;
				if(!(rs.EOF||rs.BOF)) {
					showStr += 'v' + 'ar refArray = n' + 'ew Array();\n';
					showStr += 'v' + 'ar countArray = n' + 'ew Array();\n';
					showStr += 'v' + 'ar count;\n';
					while(!(rs.EOF||rs.BOF)) {
						showStr += 'refArray[' + c + '] = \'' + new String(rs("refUrl")) + '\';\n';
						showStr += 'countArray[' + c + '] = ' + new String(rs("refCount")) + ';\n';
						rs.MoveNext();
						c++;
					}
					showStr += 'count = ' + c + ';\n';
					showStr += 'f' + 'or(v' + 'ar i = 0;i < count;i++){;\n';
					showStr += '	d' + 'ocument.write(\' - <a href = "\' + unescape(refArray[i]) + \'" target="_blank">\' + unescape(refArray[i]).substring(7,40) + \'</a>\');\n';
					showStr += '	d' + 'ocument.write(\'[\' + countArray[i] + \']<br/>\');\n';
					showStr += '};\n';
				}
				rs.close();
			}
			showStr += 'd' + 'ocument.write(\'<a href = ' + core.serverUrl + '>' + core.brandName + '</a> -- ' + core.authorInfo + '</div>\');\n';
		}
		
		//show type
		if(obShowType > 1) {
			showPR();
		}
		if(obShowType == 1 || obShowType == 3) {
			showCount();
		}
		Response.write(showStr);
	}
	
	this.getPR = function() {
		var google;
		var pagerank;
		if(obUrl != ''){
			google = getGoogleHostInfo(obUrl);
			pagerank = getPageRank(google);
		}else{
			pagerank = 0;
		}
		return pagerank;
	}

}
var ob = new Observer();
ob.initialize();
ob.showInfo();
Response.flush();
</SCRIPT>