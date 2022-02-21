<!--#include file="conn.asp"-->
<%
set rs=conn.execute("select classid,title,content,writer,writefrom,addtime from article where articleid="&request("id"))
if rs.eof and rs.bof then
response.redirect "index.asp"
response.end
else
dim classid
dim title,content,writer,writefrom,addtime
classid=rs(0)
title=rs(1)
content=rs(2)
writer=rs(3)
writefrom=rs(4)
addtime=rs(5)
end if
rs.close%>
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
          <td width="171" valign="top" bgcolor="F7EEE4"><!--#include file="left.inc"-->          </td>
          <td width="609" valign="top"><table width="580" border="0" align="center" cellpadding="5" cellspacing="5">
            <tr>
              <td width="590" colspan="2"><img src="images/deng<%=classid%>.gif" width="580" height="60"></td>
            </tr>
          </table>
            <table cellSpacing="5" cellPadding="5" width="580" align="center" border="0" id="table1">
				<tr>
					<td align="middle" width="590" colSpan="2">
					<span class="txt4"><strong><%=title%></strong><br>
　</span><hr SIZE="1"><%if writer<>"" then%> 作者： <%=writer%>　<%end if%>　来源：<%if writefrom="" then response.write("SDDYZ.com") else  response.write writefrom%>　　时间：<%=addtime%> 
					</td>
				</tr>
				<tr>
					<td colSpan="2"><span class="txt4">
					<%=content%></span></td>
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
<%set rs=nothing
            conn.close
            set conn=nothing%>