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
<title>�鿴����</title>
<link rel=stylesheet href="style.css" type="text/css">
</head>
<body>
<div align="center">
  <center>
<table border="0" width="500" cellspacing="0" cellpadding="0">
    <tr>
      <td width="100%"><% 
      set rst=server.createobject("adodb.recordset")

      sql="select * from guestbook order by id desc"

      rst.open sql,conn,3,1      
	if Not(rst.bof and rst.eof) then'�б����ݱ����Ƿ�Ϊ�ռ�¼
			NumRecord=rst.recordcount
			rst.pagesize=10
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
        <div align="center">
          <center>
        <table border="0" width="530" bordercolorlight="#000000" cellspacing="0" cellpadding="2" bordercolordark="#FFFFFF">

            <%if Not(rst.bof and rst.eof) then
	rst.move (Cint(NoncePage)-1)*10,1
	for i=1 to rst.pagesize
%>
<tr>
<td>
<table border="0" width="100%" cellspacing="0" cellpadding="5">
  <tr>
    <td width="35%" align="right" bgcolor="#F2F2F2">�û�����</td>
    <td width="65%"><%=rst("user_name")%></td>
  </tr>
  <tr>
    <td width="35%" align="right" bgcolor="#F2F2F2">���⣺</td>
    <td width="65%"><%=rst("user_title")%></td>
  </tr>
  <tr>
    <td width="35%" align="right" bgcolor="#F2F2F2">���ݣ�</td>
    <td width="65%"><%=rst("user_text")%></td>
  </tr>
  <tr>
    <td width="35%" align="right" bgcolor="#F2F2F2">����ʱ�䣺</td>
    <td width="65%"><%=rst("user_time")%></td>
  </tr>
  <tr>
    <td width="100%" align="right" bgcolor="#F2F2F2" colspan="2"><a href="del_guest.asp?id=<%=rst("id")%>">ɾ��������</a></td>
  </tr>
</table>
</td>
</tr>         
<%rst.movenext
if rst.eof then exit for
next
else
	response.write "<tr><td colspan=13><marquee scrolldelay=120 behavior=alternate>û���ҵ��κμ�¼!!!</marquee></td></tr>"
end if	
rst.close
set rst=nothing
%>
</center>
  </center>
        </table>
        </div>
  </table>
</div>
<table width="550" border="0" align="center">
  <tr> 
   <td height="17">             
      <div align="right"> 
        <input type="hidden" name="page" value="<%=NoncePage%>">
              <%
if NoncePage>1 then
	response.write "|<a href=viewguest.asp?page=1>�� ҳ</a>| |<a href=viewguest.asp?page="&NoncePage-1&">��һҳ</a>|&nbsp"
else
	response.write "|�� ҳ| |��һҳ|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=viewguest.asp?page="&NoncePage+1&">��һҳ</a>| |<a href=viewguest.asp?page="&NumPage&">β ҳ</a>|"
else
	response.write "|��һҳ| |β ҳ|"
end if
%>
 &nbsp;ҳ�Σ�<font color="#0033CC"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 
  ��<font color="#0033CC"><%=NumRecord%></font>����¼&nbsp; </div>                                          
          </td>
    </tr>
  </table>
</body>
</html>
