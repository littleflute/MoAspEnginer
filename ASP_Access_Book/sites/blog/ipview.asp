<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
    <td width="128" align="center" valign="top" nowrap bgcolor="#FFFFFF"><br>
<b>IP地址查询<br>
<br>
<a href="javascript:window.close();">点击关闭窗口</a><br>
</b></td>
    <td width="100%" align="center" valign="top" bgcolor="#FFFFFF"><%Dim ipdata
	ipdata=CheckStr(Request.QueryString("ipdata"))
	If ipdata=Empty Then ipdata=CheckStr(Request.Form("ipdata"))
	If ipdata=Empty Then
	%>
      <table width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bordercolor="#CCCCCC">
        <form name="ipsearch" method="post" action="">
        <tr class="msg_head">
          <td colspan="2">IP地址查询</td>
        </tr>
        <tr>
          <td width="35%">请输入你需要查询的IP：<br>
          [按照XXX.XXX.XXX.XXX格式]</td>
          <td width="65%"><input name="ipdata" type="text" id="ipdata" size="50"></td>
        </tr>
        <tr align="center">
          <td colspan="2">
            <input type="submit" name="Submit" value=" 查  询 ">
          </td>
        </tr></form>
      </table>
      <br>
      <%Else
	  Dim IPLocation
	  If 1=0 Then
		  IPLocation="IP地址格式不合法！"
	  Else
		  Dim ConnIP
		  Set ConnIP=Server.CreateObject("ADODB.Connection")
		  ConnIP.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(AccessPath&"/"&IPAccessFile)
		  ConnIP.Open
		  Dim IPDataArr,IPDataStr,IPQuery
		  IPDataArr=Split(ipdata,".")
		  IPDataStr=Cint(IPDataArr(0))*16777216 + Cint(IPDataArr(1))*65536 + Cint(IPDataArr(2))*256 + Cint(IPDataArr(3))
		  Set IPQuery=Server.CreateObject("ADODB.RecordSet")
		  SQL = "SELECT TOP 1 ip_Local FROM [blog_IPData] WHERE ip_Start<="&IPDataStr&" AND ip_End>="&IPDataStr
		  IPQuery.Open SQL,ConnIP,1,1
		  SQLQueryNums=SQLQueryNums+1
		  If IPQuery.EOF AND IPQuery.BOF Then
				IPLocation="库中没有这个IP记录"
		  Else
				IPLocation=IPQuery("ip_Local")
		  End IF
		  IPQuery.Close
		  SET IPQuery=Nothing
		  ConnIP.Close
		  Set ConnIP=Nothing
		End If
	  %><br>
<div class="msg_head">你所查询的IP地址：<%=ipdata%> 的来源</div><div class="msg_content"><br />
  <strong>[ <%=IPLocation%> ]</strong><br />
  <br>
</div>
<br><%End IF%></td></tr></table>
<!--#include file="footer.asp" -->