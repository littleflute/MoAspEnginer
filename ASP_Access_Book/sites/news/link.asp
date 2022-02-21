<!--#include file="conn.asp"-->
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
          <td width="171" valign="top" bgcolor="F7EEE4"><!--#include file="left.inc"-->
          </td>
          <td width="609" valign="top"><table width="580" border="0" align="center" cellpadding="5" cellspacing="5">
            <tr>
              <td width="590" colspan="2">
				<img src="images/deng17.gif" width="580" height="60"></td>
            </tr>
          </table><br><table width="560"  border="0" align="center" cellpadding="3" cellspacing="3" class="txt1">
            <%set rs=conn.execute("select title,content from article where classid=17 order by addtime desc")
            do while not rs.eof%><tr>
              <td width="3%"><img src="images/iocn1.gif" width="7" height="7"></td>
              <td width="47%" class="txt1"><a href="<%=rs(1)%>" target="_blank" class="list3"><%=rs(0)%></a></td>
              <td width="3%"><img src="images/iocn1.gif" width="7" height="7"></td>
              <td width="47%"><%rs.movenext
              if not rs.eof then %>
              <a href="<%=rs(1)%>" target="_blank" class="list3"><%=rs(0)%></a>
              <%rs.movenext
              end if%></td>
            </tr>
            <%
            loop
            rs.close
            set rs=nothing
            conn.close
            set conn=nothing%>
            <tr align="center">
              <td colspan="4">　</td>
              </tr>
          </table>          </td>
        </tr>
      </table>
      <!--#include file="end.inc"--></td>
  </tr>
</table>
</body>
</html>