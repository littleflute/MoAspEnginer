<!--#include file="conn.asp"-->
<!--#include file="fenlei.asp"-->
<%
id = request("id")
if id = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('没有找到您要查看的记录');window.close();</script>"
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
 response.write "<script>alert('没有找到您要查看的记录');window.close();</script>"
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
<title>详细资料</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<br>
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=5 width="350" border=1>
<tr><td>
<table align=center width=100% border=0>
<tr><td align="center" class="header" colspan=2>详细资料</td></tr>
<tr><td width=80 align="right"><b>所属<%=strfenlei1%>：</b></td><td width=270><%=rs("fenlei1")%></td></tr>
<tr><td align="right"><b>所属<%=strfenlei2%>：</b></td><td><%=rs("fenlei2")%></td></tr>
<tr><td align="right"><b>教师姓名：</b></td><td><%=rs("teacher")%></td></tr>
<tr><td align="right"><b>课程名称：</b></td><td><%=rs("course")%></td></tr>
<%
dateandtime = rs("dateandtime")
%>
<tr><td align="right"><b>更新日期：</b></td><td><%=year(dateandtime)%>年<%=month(dateandtime)%>月<%=day(dateandtime)%>日</td></tr>
<tr><td align="right"><b>更新时间：</b></td><td><%=hour(dateandtime)%>:<%=minute(dateandtime)%>:<%=second(dateandtime)%></td></tr>
<tr><td align="right"><b>阅读次数：</b></td><td><%=rs("times")%></td></tr>
<tr><td align="right"><b>资料大小：</b></td><td><%=rs("filesize")%><b>K</b></td></tr>
<tr><td align="right"><b>资料类型：</b></td><td><%=rs("type")%></td></tr>
<tr><td align="right"><b>资料标题：</b></td><td><%=rs("title")%></td></tr>
<tr><td align="right" valign="top"><b>资料简介：</b></td><td rowspan=2 valign="top"><%=rs("content")%>&nbsp;</td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td colspan=2 bgColor=#cccccc height=1></td></tr>
<tr>
          <td colspan=2 align=center> <br>
            <img src=images/close.gif border=0><a href=# onclick=javascript:window.close();>关闭窗口</a>&nbsp;&nbsp; 
            <img src=images/download.gif border=0><a href=download.asp?id=<%=id%> target=_blank>点这里阅读</a> 
            <img src=images/close.gif border=0><a href=redetail.asp?id=<%=id%> target=_blank>提交作业</a></td>
        </tr>
<tr><td colspan=2 bgColor=#cccccc height=1></td></tr>
<tr><td colspan=2 align=center>
<a href=# onclick="window.open('score.asp?id=<%=id%>','score','width=350,height=200,resizable=no,scrollbars=yes,menubar=no,status=no');" title="点击查看当前得分情况">
<font color=green>您的评价</font></a>&nbsp;&nbsp;
很差<input type=radio name="score" onclick="window.open('score.asp?id=<%=id%>&score=1','score','width=350,height=250,resizable=no,scrollbars=yes,menubar=no,status=no');">
较差<input type=radio name="score" onclick="window.open('score.asp?id=<%=id%>&score=2','score','width=350,height=250,resizable=no,scrollbars=yes,menubar=no,status=no');">
一般<input type=radio name="score" onclick="window.open('score.asp?id=<%=id%>&score=3','score','width=350,height=250,resizable=no,scrollbars=yes,menubar=no,status=no');">
良好<input type=radio name="score" onclick="window.open('score.asp?id=<%=id%>&score=4','score','width=350,height=250,resizable=no,scrollbars=yes,menubar=no,status=no');">
优秀<input type=radio name="score" onclick="window.open('score.asp?id=<%=id%>&score=5','score','width=350,height=250,resizable=no,scrollbars=yes,menubar=no,status=no');">
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