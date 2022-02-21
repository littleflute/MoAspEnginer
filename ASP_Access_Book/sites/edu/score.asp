<!--#include file="conn.asp"-->
<%
id = trim(request("id"))
score = trim(request("score"))
if IsNumeric(score) then
 score = cint(score)
else
 score = 0
end if
if id = "" or score < 1 or score > 5 then
 flag = 0
else
 if request.cookies("scored"&id) = "yes" then
  flag = 1
 else
  response.cookies("scored"&id) = "yes"
  response.cookies("scored"&id).expires = now() + 1/72
  flag = 2
 end if
end if
sql = "select * from main where mainid="&id
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('请不要捣乱');window.close();</script>"
 response.end
else
 rs.close
 if flag=2 then conn.execute "insert into score (idofmain,score) values("&id&","&score&")"
end if

dim curscore(5)
for i = 1 to 5
 sql = "select count(scoreid) from score where score="&i&" and idofmain="&id
 rs.open sql,conn,1,1
 curscore(i-1) = rs(0)
 rs.close
next
set rs = nothing
conn.close
set conn = nothing

maxscore = curscore(0)
for i = 1 to 4
 if curscore(i) > maxscore then maxscore = curscore(i)
next
if maxscore < 1 then maxscore = 1
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
<title>得票情况</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<br>
<%if flag=1 then response.write "<center><font color=red>请不要重复投票</font></center><br>"%>
<%if flag=2 then response.write "<center><font color=red>非常感谢您的意见</font></center><br>"%>
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=5 width="300" border=1>
<tr align=center class="category"><td colspan=2>当前得票情况</td></tr>
<tr align=center><td align=center width=50>优秀</td>
<td align=left width=250><img src=images/bar1.gif width=<%=cint(curscore(4)/maxscore*200)%> height=10>（<%=curscore(4)%>）</td></tr>
<tr align=center><td align=center width=50>良好</td>
<td align=left width=250><img src=images/bar2.gif width=<%=cint(curscore(3)/maxscore*200)%> height=10>（<%=curscore(3)%>）</td></tr>
<tr align=center><td align=center width=50>一般</td>
<td align=left width=250><img src=images/bar3.gif width=<%=cint(curscore(2)/maxscore*200)%> height=10>（<%=curscore(2)%>）</td></tr>
<tr align=center><td align=center width=50>较差</td>
<td align=left width=250><img src=images/bar4.gif width=<%=cint(curscore(1)/maxscore*200)%> height=10>（<%=curscore(1)%>）</td></tr>
<tr align=center><td align=center width=50>很差</td>
<td align=left width=250><img src=images/bar5.gif width=<%=cint(curscore(0)/maxscore*200)%> height=10>（<%=curscore(0)%>）</td></tr>
</table>
</body>
</html>