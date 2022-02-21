<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn.asp"-->
<table>
<%
sql_news="select  top 5  * from bbs where layer=1 order by submit_date desc"
set rs_news=conn.execute(sql_news)
if not rs_news.eof then
	rs_news.movefirst
	do while not rs_news.eof
	
%> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
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
  <tr bgcolor="#FFFFCC" align="center">
					<td><a href="cyz/count_hits.asp?bbs_id=<%=rs("bbs_ID")%>"  target="_blank"><%=RS("title")%></a></td>
					<td>[<%=RS("child")%>/<%=RS("hits")%>]</td>
					<td>[<%=rs("user_name")%>]</td>
					<td>[<%=rs("submit_date")%>]</td>
				</tr>			
	<%	rs_news.movenext
	loop
	end if
rs_news.close
		</table>
		
</body>
</html>

