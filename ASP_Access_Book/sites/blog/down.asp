<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/ubbcode.asp" -->
<%
Dim downID
downID=CheckStr(Trim(Request.QueryString("downID")))
If Not IsInteger(downID) Then downID=0
IF downID=Empty Then
	Response.Write("<br>请选择一个下载对象。<a href='default.asp'>返回主页</a>")
Else
	Dim downLog
	Set downLog=Server.CreateObject("ADODB.RecordSet")
	SQL="SELECT downl_ID,downl_Dcomm1URL,downl_Dcomm2URL,downl_Dcomm3URL,downl_Dcomm4URL,downl_Nums FROM blog_Download WHERE downl_ID="&downID&""
	downLog.Open SQL,Conn,1,3
Dim downl_Nums,downURL1,downURL2,downURL3,downURL4
downURL1=downLog("downl_Dcomm1URL")
downURL2=downLog("downl_Dcomm2URL")
downURL3=downLog("downl_Dcomm3URL")
downURL4=downLog("downl_Dcomm4URL")
downl_Nums=downLog("downl_Nums")
	if Request.QueryString("action")="Url_1" then
	Response.Write("<meta http-equiv='refresh' content='0;URL="&downURL1&"'>")
    elseif Request.QueryString("action")="Url_2" then
	Response.Write("<meta http-equiv='refresh' content='0;URL="&downURL2&"'>")
    elseif Request.QueryString("action")="Url_3" then
	Response.Write("<meta http-equiv='refresh' content='0;URL="&downURL3&"'>")
    elseif Request.QueryString("action")="Url_4" then
	Response.Write("<meta http-equiv='refresh' content='0;URL="&downURL4&"'>")
	end if
Conn.ExeCute("UPDATE blog_Download SET downl_Nums=downl_Nums+1 WHERE downl_ID="&downID&"")
downLog.UPDATE
downLog.Close
end if%>
