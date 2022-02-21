<!--#include file="conn.asp"-->
<!--#include file="code.asp"-->
<!--#include file="char.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>论坛</title>
<style type="text/css">
<!--
a:visited{text-transform: none; text-decoration: none; font-style: normal; font-weight: normal; font-variant: normal; color: #000000}
a:hover{text-decoration:underline; color: #FF3333}
a:link{text-transform: none; text-decoration: none; font-weight: normal; font-variant: normal; color: #000000}
.p9 { font-size: 9pt; line-height: 13pt; text-decoration: none}
.pic_show { filter:dropshadow(color=#999999,offx=5,offy=5.positive=1)}
.ptitle_9 { font-size: 9pt; line-height: 11pt}
.pbutton { border: #666666; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
-->
</style>
</head>

<body>
<table class="p9">
  <tr border="0" cellspacing="0" cellpadding="0"><td>
<%
maxLen=45
function gotTopic(str,strlen)
	dim l,t,c, i
	l=len(str)
	t=0
	for i=1 to l
	c=Abs(Asc(Mid(str,i,1)))
	if c>255 then
	t=t+2
	else
	t=t+1
	end if
	if t>=strlen then
	gotTopic=left(str,i)&"..."
	exit for
	else
	gotTopic=str
	end if
	next
end function
sql="select top 8 * from news where layer=1 order by bbs_id desc"
set rs=server.CreateObject ("ADODB.RecordSet")
	rs.open sql,conn,1,1
	do while not rs.eof
	
topic=gotTopic(ubbcode(rs("title")),maxLen)
%>
·<a href="count_hits.asp?bbs_id=<%=rs("bbs_ID")%>&boardid=<%=rs("boardid")%>" target="_blank"><%=(topic) %></a><br>
<%
	rs.movenext
	loop
	rs.close
%>
</td>
</tr>	
</body>
</html>
