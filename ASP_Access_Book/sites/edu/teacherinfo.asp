<!--#include file="conn.asp"-->
<!--#include file="fenlei.asp"-->
<%
id = trim(request("id"))

if id = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('�벻Ҫ����');top.window.location.href='index.asp';</script>"
 response.end
end if

if session("admin") = "admin" then
 isadmin = true
else
 isadmin = false
end if

if session("teacherid") <> "" then
 isteacher = true
else
 isteacher = false
end if

sql = "select * from teacher where teacherid="&id
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE>TD {
	FONT-SIZE: 9pt; LINE-HEIGHT: 140%
}
BODY {
	FONT-SIZE: 9pt; LINE-HEIGHT: 140%
}
A:link {
	COLOR: #0033cc; TEXT-DECORATION: none
}
A:visited {
	COLOR: #0033cc; TEXT-DECORATION: none
}
A:active {
	COLOR: #ff0000; TEXT-DECORATION: none
}
A:hover {
	COLOR: #000000; TEXT-DECORATION: underline
}
.header	{
font-family: Tahoma, Verdana; font-size: 9pt; color: #FFFFFF; background-color: #00CCCC
}
.category{
font-family: Tahoma, Verdana; font-size: 9pt; color: #000000; background-color: #EFEFEF
}
</STYLE>
<title>����ר��</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<%
if not isadmin and not isteacher then
%>
<!--#include file="head.asp"-->
<%
end if
if rs.bof and rs.eof then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<center>û���ҵ���Ҫ�鿴�Ľ�ʦ</center>"
 response.end
else
%>
<br>
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=1 width="740" border=1>
<tr><td align="center" class="header" colspan=3>��������</td></tr>
<tr><td width=80 align="right"><b>����<%=strfenlei1%>��</b></td><td width=420>&nbsp;<%=rs("fenlei1")%></td>
    <td align="center" valign="middle" rowspan=7 width=240>
<%
if left(rs("photourl"),6) = "files/" then
 response.write "<a href="&rs("photourl")&" target=_blank><img src="&rs("photourl")&" border=0 height=150 width=200></a>"
else
 response.write "������Ƭ"
end if
%>
    </td>
</tr>
<tr><td width=80 align="right"><b>����<%=strfenlei2%>��</b></td><td width=420>&nbsp;<%=rs("fenlei2")%></td></tr>
<tr><td width=80 align="right"><b>��ʦ������</b></td><td width=420>&nbsp;<%=rs("teacher")%></td></tr>
<tr><td width=80 align="right"><b>E-Mail��</b></td><td width=420>&nbsp;<a href=mailto:<%=rs("email")%>><%=rs("email")%></a></td></tr>
<tr><td width=80 align="right"><b>������ҳ��</b></td><td width=420>&nbsp;<a href=<%=rs("homepage")%> target=_blank><%=rs("homepage")%></a></td></tr>
<tr><td width=80 align="right"><b>QQ���룺</b></td><td width=420>&nbsp;<%=rs("qq")%></td></tr>
<tr><td width=80 align="right"><b>ͨѶ��ַ��</b></td><td width=420>&nbsp;<%=rs("address")%></td></tr>
<tr><td width=80 align="right" valign="top"><b>���˼�飺</b></td><td valign="top" colspan=2>&nbsp;&nbsp;<%=rs("intro")%>&nbsp;</td></tr>
<%
rs.close
end if
sql = "select * from type"
rs.open sql,conn,1,1
set rs1 = server.createobject("adodb.recordset")
do while not rs.eof
 sql1 = "select count(mainid) from main where idofteacher="&id&" and idoftype="&rs("typeid")
 rs1.open sql1,conn,1,1
 counter = rs1(0)
 rs1.close
 if counter mod 3 = 0 then
  iframeheight = 20*(int(counter/3)+1)
 else
  iframeheight = 20*(int(counter/3)+2)
 end if
 if iframeheight > 220 then iframeheight = 220
%>
<tr><td align="center" class="header" colspan=3>���ڱ�վ������<%=rs("type")%>����<%=counter%>����</td></tr>
<tr><td align="center" colspan=3>
<iframe name="titleof<%=rs("typeid")%>" frameborder=0 width=100% height=<%=iframeheight%> scrolling=no src=titlelist.asp?type=<%=rs("typeid")%>&id=<%=id%>></iframe>
</td></tr>
<%
 rs.movenext
loop
set rs1 = nothing
rs.close
set rs = nothing
%>
</table>
<%
if not isadmin and not isteacher then
%>
<!--#include file="foot.asp"-->
<%
end if
%>
</body>
</html>
<%
conn.close
set conn = nothing
%>