<!--#include file="conn.asp"-->
<!--#include file="fenlei.asp"-->
<%
id = request("id")
if id = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('û���ҵ���Ҫ�鿴�ļ�¼');window.close();</script>"
 response.end
end if

set rs = server.createobject("adodb.recordset")
sql = "select * from main,teacher,type where main.idofteacher=teacher.teacherid and main.idoftype=type.typeid and main.mainid="&id
rs.open sql,conn,1,1
if rs.bof and rs.eof then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('û���ҵ���Ҫ�鿴�ļ�¼');window.close();</script>"
 response.end
else
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
<title>��ϸ����</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<br>
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=5 width="350" border=1>
<tr><td>
<table align=center width=100% border=0>
<tr><td align="center" class="header" colspan=2>��ϸ����</td></tr>
<tr><td width=80 align="right"><b>����<%=strfenlei1%>��</b></td><td width=270><%=rs("fenlei1")%></td></tr>
<tr><td align="right"><b>����<%=strfenlei2%>��</b></td><td><%=rs("fenlei2")%></td></tr>
<tr><td align="right"><b>��ʦ������</b></td><td><%=rs("teacher")%></td></tr>
<tr><td align="right"><b>�γ����ƣ�</b></td><td><%=rs("course")%></td></tr>
<%
dateandtime = rs("dateandtime")
%>
<tr><td align="right"><b>�������ڣ�</b></td><td><%=year(dateandtime)%>��<%=month(dateandtime)%>��<%=day(dateandtime)%>��</td></tr>
<tr><td align="right"><b>����ʱ�䣺</b></td><td><%=hour(dateandtime)%>:<%=minute(dateandtime)%>:<%=second(dateandtime)%></td></tr>
<tr><td align="right"><b>�Ķ�������</b></td><td><%=rs("times")%></td></tr>
<tr><td align="right"><b>���ϴ�С��</b></td><td><%=rs("filesize")%><b>K</b></td></tr>
<tr><td align="right"><b>�������ͣ�</b></td><td><%=rs("type")%></td></tr>
<tr><td align="right"><b>���ϱ��⣺</b></td><td><%=rs("title")%></td></tr>
<tr><td align="right" valign="top"><b>���ϼ�飺</b></td><td rowspan=2 valign="top"><%=rs("content")%>&nbsp;</td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td colspan=2 bgColor=#cccccc height=1></td></tr>
<tr>
          <td colspan=2 align=center> <br>
            <img src=images/close.gif border=0><a href=# onclick=javascript:window.close();>�رմ���</a>&nbsp;&nbsp; 
            <img src=images/download.gif border=0><a href=download.asp?id=<%=id%> target=_blank>�������Ķ�</a> 
            <img src=images/close.gif border=0><a href=redetail.asp?id=<%=id%> target=_blank>�ύ��ҵ</a></td>
        </tr>
<tr><td colspan=2 bgColor=#cccccc height=1></td></tr>
<tr><td colspan=2 align=center>
<a href=# onclick="window.open('score.asp?id=<%=id%>','score','width=350,height=200,resizable=no,scrollbars=yes,menubar=no,status=no');" title="����鿴��ǰ�÷����">
<font color=green>��������</font></a>&nbsp;&nbsp;
�ܲ�<input type=radio name="score" onclick="window.open('score.asp?id=<%=id%>&score=1','score','width=350,height=250,resizable=no,scrollbars=yes,menubar=no,status=no');">
�ϲ�<input type=radio name="score" onclick="window.open('score.asp?id=<%=id%>&score=2','score','width=350,height=250,resizable=no,scrollbars=yes,menubar=no,status=no');">
һ��<input type=radio name="score" onclick="window.open('score.asp?id=<%=id%>&score=3','score','width=350,height=250,resizable=no,scrollbars=yes,menubar=no,status=no');">
����<input type=radio name="score" onclick="window.open('score.asp?id=<%=id%>&score=4','score','width=350,height=250,resizable=no,scrollbars=yes,menubar=no,status=no');">
����<input type=radio name="score" onclick="window.open('score.asp?id=<%=id%>&score=5','score','width=350,height=250,resizable=no,scrollbars=yes,menubar=no,status=no');">
</td></tr>
</table>
</td></tr></table>
<br>
</body>
</html>
<%
rs.close
set rs = nothing
conn.close
set conn = nothing
end if
%>