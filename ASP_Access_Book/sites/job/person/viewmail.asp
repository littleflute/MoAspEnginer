<% Response.Buffer=True %>
<!--#include file="../inc/person.inc"-->
<% uname=session("puid") 
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from pmailbox where reid='"&uname&"'and newmail=0"
rs.open sql,conn,1,1
mailnum=rs.recordcount
rs.close                                                                    
set rs=nothing
Set rs = Server.CreateObject("ADODB.Recordset")
if request("del")<>"" then 
conn.Execute("delete * from pmailbox where id="&request("del"))
response.write"<SCRIPT language=JavaScript>alert('�ż�ɾ���ɹ���');"
response.write"javascript:window.close();</SCRIPT>" 
end if
sql="select * from pmailbox where id="&request("id")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
response.write"<SCRIPT language=JavaScript>alert('�ż������ڻ��ѱ�ɾ����');"
response.write"javascript:window.close();</SCRIPT>"
end if
newmail=rs("newmail") %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" href="../inc/register.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=rs("title")%></title>
</head>

<body>
<div align="center">
  <center>
  <table border="4" cellpadding="0" cellspacing="0" width="366" height="69" bordercolor="#EBEEF3" bordercolorlight="#EBEEF3" bordercolordark="#EBEEF3">
    <tr>
      <td height="18" width="356" valign="bottom" bgcolor="#C6CEDE" colspan="2" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
        <p align="center">=== վ���ż� ===</td>              
    </tr>
    <tr>
      <td height="20" width="356" valign="bottom" bgcolor="#EBEEF3" colspan="2">&nbsp;������:<%=rs("sendname")%></td>
    </tr>
    <tr>
      <td height="19" width="356" valign="bottom" bgcolor="#EBEEF3" colspan="2">&nbsp;��&nbsp;             
        ��:<%=rs("title")%></td>
    </tr>
    <tr>
      <td height="19" width="356" valign="bottom" bgcolor="#EBEEF3" colspan="2">&nbsp;����ʱ��:<%=rs("sdate")%></td>
    </tr>
    <tr>
      <td width="2" rowspan="3">
  </center>
    </td>
      <td height="4" width="349">
    </td>
    </tr>
  <tr>
      <td bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" width="349">
	 &nbsp;&nbsp;&nbsp;&nbsp;<%=rs("mailtext")%>
    </td>
  </tr>
  <tr>
      <td height="1" width="349">
    </td>
  </tr>
  <tr>
      <td height="23" width="356" valign="bottom" bgcolor="#EBEEF3" colspan="2">
  <p align="center">[<a href="viewmail.asp?del=<%=rs("id")%>">ɾ���ż�</a>] [<a href="sendmail.asp?reid=<%=rs("senduid")%>">�ظ��ż�</a>]         
    </td>
  </tr>
  </table>
</div>
<p align="center">
��<a href="javascript:window.close()">�رմ���</a>��

</body>
</html>
<% rs.close                                                                    
set rs=nothing
if newmail="0" then
Set rs = Server.CreateObject("ADODB.Recordset")
sql2="update pmailbox set newmail=newmail+1 where id="&request("id")
rs.open sql2,conn,1,1
end if %>