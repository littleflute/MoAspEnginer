<!--#include file="conn.asp"-->
<%
set rs=conn.execute("select classid,title,content,picurl,writefrom,addtime from article where articleid="&request("id"))
if rs.eof and rs.bof then
response.redirect "index.asp"
response.end
else
dim classid
dim title,content,picurl,writefrom,addtime
classid=rs(0)
title=rs(1)
content=rs(2)
picurl=rs(3)
writefrom=rs(4)
addtime=rs(5)
end if
rs.close
if classid=20 then
classid=2
end if%>
<html>
<head>
<title>新闻网</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="images/css.css" rel=stylesheet type=text/css>
<style type="text/css">
<!--
body {
	background-image: url(images/pagebg1.gif);
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
</head>
<body>
<table width="780" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CACACA">
  <tr>
    <td bgcolor="#FFFFFF"><!--#include file="top.inc"-->
    <table width="780" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="171" valign="top" bgcolor="F7EEE4"><!--#include file="left.inc"-->           </td>
          <td width="609" valign="top"><table width="580" border="0" align="center" cellpadding="5" cellspacing="5">
            <tr>
              <td width="590" colspan="2"><img src="images/deng<%=classid%>.gif" width="580" height="60"></td>
            </tr>
          </table>            <br>            <table width="570"  border="0" align="center" cellpadding="4" cellspacing="4">
            <tr>
                <td class="txt5"><%=title%></td>
            </tr>
            <tr>
              <td class="txt1"><%if trim(picurl)<>"" then response.write "<img src='"&picurl&"' width='120' hspace='6' vspace='3' border='0' align='left'>" end if%><%=content%></td>
            </tr>
            </table>
</td>
        </tr>
      </table>
      <!--#include file="end.inc"--></td>
  </tr>
</table>
</body>
</html>
