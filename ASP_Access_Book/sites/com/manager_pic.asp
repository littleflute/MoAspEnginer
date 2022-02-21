<!--#include file="conn.asp"-->
<%
if session("admin")="" then
	response.redirect "admin.asp"
	response.end
end if
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>管理图片</title>
<link rel=stylesheet href="style.css" type="text/css">
<style type="text/css">
<!--
body {
	background-color: #5295C9;
}
-->
</style></head>

<body>
<div align="center">
  <center>
    <table border="0" width="755" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="02.jpg" width="756" height="30">
      <tr> 
        <td width="100%"> 
          <% 
      set rst=server.createobject("adodb.recordset")
      sql="select * from pic order by id desc"
      rst.open sql,conn,3,1      
	if Not(rst.bof and rst.eof) then'判别数据表中是否为空记录
			NumRecord=rst.recordcount
			rst.pagesize=30
			NumPage=rst.Pagecount
			if request("page")=empty then 
			NoncePage=1
		else
		if Cint(request("page"))<1 then
			NoncePage=1
		else
			NoncePage=request("page")
		end if
		if Cint(Trim(request("page")))>Cint(NumPage) then NoncePage=NumPage
	end if
else
	NumRecord=0
	NumPage=0
	NoncePage=0
	end if
%>
          <table border="0" width="754" bordercolorlight="#000000" cellspacing="1" cellpadding="2" bordercolordark="#FFFFFF">
            <tr bgcolor="#00CCCC" background="titl2.gif"> 
              <td width="71" align="center">图片名称</td>
              <td width="71" align="center">图片类别</td>
              <td width="72" align="center">图片型号</td>
              <td width="71" align="center">图片说明</td>
              <td width="72" align="center">发表时间</td>
              <td width="71" align="center">删除</td>
            </tr>
            <%if Not(rst.bof and rst.eof) then
	rst.move (Cint(NoncePage)-1)*30,1
	for i=1 to rst.pagesize
%>
            <tr bgcolor="#A2DFEA"> 
              <td width="71" align="center"><%=rst("pic_name")%></td>
              <td width="71" align="center"><%=rst("pic_class")%></td>
              <td width="72" align="center"><%=rst("pic_type")%></td>
              <td width="72" align="center"><%=rst("pic_text")%></td>
              <td width="72" align="center"><%=rst("pic_time")%></td>
              <td width="72" align="center"><a href=del_pic.asp?id=<%=rst("id")%>>删除</a></td>
            </tr>
            <%		 		rst.movenext
			   		if rst.eof then exit for
					next
else
	response.write "<tr><td colspan=13><marquee scrolldelay=120 behavior=alternate>没有找到任何记录!!!</marquee></td></tr>"
end if	

rst.close
set rst=nothing

%>
          </table>
    </table>
  </center>
</div>
<table width="754" border="0" align="center" bgcolor="#CCCCCC">
  <tr> 
          
    <td height="17" bgcolor="#0099CC"> 
      <div align="right"> 
        <input type="hidden" name="page" value="<%=NoncePage%>">
              <%
if NoncePage>1 then
	response.write "|<a href=manager_pic.asp?page=1>首 页</a>| |<a href=manager_pic.asp?page="&NoncePage-1&">上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=manager_pic.asp?page="&NoncePage+1&">下一页</a>| |<a href=manager_pic.asp?page="&NumPage&">尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
              &nbsp;页次：<font color="#0033CC"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 
              共<font color="#0033CC"><%=NumRecord%></font>条记录&nbsp; </div>                                        
          </td>
    </tr>
  </table>

</body>

</html>
