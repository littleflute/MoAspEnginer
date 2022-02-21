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
              <td width="590" colspan="2"><img src="images/deng12.gif" width="580" height="60"></td>
            </tr>
          </table>            <br>           <table width="560"  border="0" align="center" cellpadding="3" cellspacing="3" class="txt1">
            <%set rs=conn.execute("select top 8 articleid,title,picurl from article where classid=12 and writefrom='' order by addtime desc")
                      i=0
                      do while not rs.eof and i<8
                      if i mod 4=0 then%>
                            <tr>
                            <%end if%>
              <td><table width="98%"  border="0" cellspacing="4" cellpadding="4">
                  <tr>
                    <td align="center" class="txt1"><a href="renwu.asp?id=<%=rs(0)%>"><img src="<%if trim(rs(2))<>"" then response.write rs(2) else response.write "images/nopicture.gif" end if%>" width="80" height="101" border="0"></a> </td>
                  </tr>
                  <tr>
                    <td align="center"><a href="renwu.asp?id=<%=rs(0)%>" class="list5"><%=rs(1)%></a></td>
                  </tr>
              </table></td>
              <%i=i+1
              if i mod 4 =0 or rs.eof then%>
              </tr><%end if
              rs.movenext
              loop
              rs.close%>
              <tr align="center">
              <td colspan="4">
				<p align="right"><a href="list2.asp?classid=12&txt=" class="list5">更多 &gt;&gt;&gt;</a></td>
            </tr>
            </table> <table width="570"  border="0" align="center" cellpadding="4" cellspacing="4" id="table1">
            <tr>
              <td class="txt5"> 特约研究员: </td>
            </tr>
          </table>           
			<table width="560"  border="0" align="center" cellpadding="3" cellspacing="3" class="txt1" id="table2">
            <%set rs=conn.execute("select top 8 articleid,title,picurl from article where classid=12 and writefrom='特约研究员' order by addtime desc")
                      i=0
                      do while not rs.eof and i<8
                      if i mod 4=0 then%>
                            <tr>
                            <%end if%>
              <td><table width="98%"  border="0" cellspacing="4" cellpadding="4">
                  <tr>
                    <td align="center" class="txt1"><a href="renwu.asp?id=<%=rs(0)%>"><img src="<%if trim(rs(2))<>"" then response.write rs(2) else response.write "images/nopicture.gif" end if%>" width="80" height="101" border="0"></a> </td>
                  </tr>
                  <tr>
                    <td align="center"><a href="renwu.asp?id=<%=rs(0)%>" class="list5"><%=rs(1)%></a></td>
                  </tr>
              </table></td>
              <%i=i+1
              if i mod 4 =0 or rs.eof then%>
              </tr><%end if
              rs.movenext
              loop
              rs.close%>
               <tr align="center">
              <td colspan="4">
				<p align="right"><a href="list2.asp?classid=12&txt=特约研究员" class="list5">更多 &gt;&gt;&gt;</a></td>
            </tr></table></td>
            </tr>
            </table><!--#include file="end.inc"--></td>
        </tr>
      </table>
</body>
</html>
