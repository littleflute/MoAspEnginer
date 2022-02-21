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
          <td width="171" valign="top" bgcolor="F7EEE4">
				<table border="0" style="border-collapse: collapse" width="100%" cellpadding="0">
					<tr>
						<td height=20 align=center><script language="JavaScript"><!-- 
var dayNames = new Array("星期日","星期一","星期二","星期三","星期四","星期五","星期六"); 
Stamp = new Date(); 
document.write("" + Stamp.getYear()+"."+(Stamp.getMonth()+ 1)+"."+Stamp.getDate()+"."+ dayNames[Stamp.getDay()] +""); 
//--> 
</script> </td>
			</tr>
		</table><!--#include file="left.inc"--> 
          <table width="171" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><img src="images/zhuan.gif" width="171" height="28" border="0" usemap="#Map8"></td>
            </tr>
            <tr>
              <td><table width="98%"  border="0" cellspacing="2" cellpadding="0">
              <%set rs=conn.execute("select top 6 articleid,title,picurl from article where classid=8 order by addtime desc")
              i=0
              do while not rs.eof and i<6
              if i mod 2 =0 then%>
                  <tr>
               <%end if%>
                    <td valign="bottom" class="txt1" align="center"><img src="<%if trim(rs(2))<>"" then response.write rs(2) else response.write "images/nopicture.gif" end if%>" width="80" height="101"><br><a  target=_blank  href="renwu.asp?id=<%=rs(0)%>" class="list5" target=_blank><%=rs(1)%></a></td>
                <%i=i+1
                if i mod 2 =0 or rs.eof then%>
                </tr>
                <%end if
                rs.movenext
                loop
                rs.close%>
                  </table></td>
            </tr>
          </table>
          
          <table width="171" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><img src="images/xuejie.gif" width="171" height="28" border="0" usemap="#Map10"></td>
            </tr>
            <tr>
              <td background="images/bg1.jpg"><table width="94%"  border="0" align="center" cellpadding="0" cellspacing="2">
                <%set rs=conn.execute("select top 4 classid,class from class where parentid=11 order by rootid,ordersid")
                      i=0
                      do while not rs.eof and i<4%>
                        <tr>
                          <td class="txt1">·<a  target=_blank  href="list.asp?classid=<%=rs(0)%>" class="list4" title="<%=rs(1)%>"><%=gotTopic(rs(1),40)%></a></td>
                        </tr>
                       <%
                       rs.movenext
			i=i+1
                       loop
                       rs.close%>

              </table></td>
            </tr>
          </table></td>
          <td width="609" valign="top"><table width="594" border="0" align="center" cellpadding="0" cellspacing="3">
            <tr>
              <td colspan="2"><img src="images/dong.gif" width="594" height="25" border="0" usemap="#Map2"></td>
            </tr>
            <tr>
            <%dim articleid,articleidpic
            set rs=conn.execute("select top 1 articleid,picurl,title from article,class where class.classid=article.classid and class.parentid=3 and picurl<>'' order by addtime desc")
            if not rs.eof then
            articleid=rs(0)
			articleidpic=articleid%>
              <td width="190"><table width="100%"  border="0" cellspacing="3" cellpadding="0">
                <tr>
                  <td><a href="show.asp?id=<%=rs(0)%>" target=_blank><img src="<%=rs(1)%>" width="180" height="133" border="0"></a></td>
                </tr>
                <tr>
                  <td align="center"> <a href="show.asp?id=<%=rs(0)%>" class="list2"><strong><%=rs(2)%></strong></a></td>
                </tr>
              </table>
              <%end if 
              rs.close%></td>
              <td width="401"><table width="100%"  border="0" cellpadding="0" cellspacing="3" class="txt1">
              <%set rs=conn.execute("select top 7 articleid,title,addtime from article where  articleid<>"&articleid&" and (classid=22 or classid=21) order by addtime desc")
              i=0
              do while not rs.eof and i<7%>
                <tr>
                  <td width="4%"><img src="images/iocn1.gif" width="7" height="7"></td>
                  <td width="77%" class="txt1"><a href="show.asp?id=<%=rs(0)%>" title="<%=rs(1)%>" class="list1" target=_blank><%=gotTopic(rs(1),40)%></a><%if month(rs(2))=month(date()) and day(rs(2))-day(date())>3 then%><img src="images/news.gif" width="30" height="13" align="absmiddle"><%end if%></td>
                  <td width="19%" class="txt3">[<%=month(rs(2))&"-"&day(rs(2))%>]</td>
                </tr>
               <%rs.movenext
               i=i+1
               loop
               rs.close%>              </table></td>
            </tr>
          </table>
          <table width="600" border="0" align="right" cellpadding="0" cellspacing="0">
            <tr>
              <td width="406" valign="top">                <table width="400" border="0" align="center" cellpadding="0" cellspacing="3">
                  <tr>
                    <td><img src="images/deng.gif" width="400" height="25" border="0" usemap="#Map"></td>
                  </tr>
                  <tr>
                    <td><table width="100%"  border="0" cellpadding="0" cellspacing="3" class="txt1">
                      <%set rs=conn.execute("select top 5 articleid,title,addtime from article where classid=9 order by addtime desc")
              		i=0
              		do while not rs.eof and i<5%><tr>
                        <td width="4%"><img src="images/iocn1.gif" width="7" height="7"></td>
                        <td width="77%" class="txt1"><a href="show.asp?id=<%=rs(0)%>" title="<%=rs(1)%>" class="list6" target=_blank><%=gotTopic(rs(1),40)%></a><%if month(rs(2))=month(date()) and day(rs(2))-day(date())>3 then%><img src="images/news.gif" width="30" height="13" align="absmiddle"><%end if%></td>
                        <td width="19%" class="txt3">[<%=month(rs(2))&"-"&day(rs(2))%>]</td>
                      </tr>
                      <tr>
                        <td background="images/iocn2.gif" height="1" colspan="3"></td>
                      </tr>
                      <%rs.movenext
               			i=i+1
               			loop
               			rs.close%> 
                    </table></td>
                  </tr>
                </table>
                <br>
                <table width="400" border="0" align="center" cellpadding="0" cellspacing="3">
                  <tr>
                    <td><img src="images/san.gif" width="400" height="25" border="0" usemap="#Map3"></td>
                  </tr>
                  <tr>
                    <td><table width="100%"  border="0" cellpadding="0" cellspacing="3" class="txt1">
                      <%set rs=conn.execute("select top 5 articleid,title,addtime from article where classid=10 order by addtime desc")
              		i=0
              		do while not rs.eof and i<5%><tr>
                        <td width="4%"><img src="images/iocn1.gif" width="7" height="7"></td>
                        <td width="77%" class="txt1"><a href="show.asp?id=<%=rs(0)%>" title="<%=rs(1)%>" class="list6" target=_blank><%=gotTopic(rs(1),40)%></a><%if month(rs(2))=month(date()) and day(rs(2))-day(date())>3 then%><img src="images/news.gif" width="30" height="13" align="absmiddle"><%end if%></td>
                        <td width="19%" class="txt3">[<%=month(rs(2))&"-"&day(rs(2))%>]</td>
                      </tr>
                      <tr>
                        <td background="images/iocn2.gif" height="1" colspan="3"></td>
                      </tr>
                      <%rs.movenext
               			i=i+1
               			loop
               			rs.close%> 
                    </table></td>
                  </tr>
                </table>
                <br>
                <table width="400" border="0" align="center" cellpadding="0" cellspacing="3">
                  <tr>
                    <td><img src="images/gong.gif" width="400" height="25" border="0" usemap="#Map3Map"></td>
                  </tr>
                  <tr>
                    <td><table width="100%"  border="0" cellpadding="0" cellspacing="3" class="txt1">
                      <%set rs=conn.execute("select top 6 articleid,title,addtime from article where classid=5 order by addtime desc")
              		i=0
              		do while not rs.eof and i<6%><tr>
                        <td width="4%"><img src="images/iocn1.gif" width="7" height="7"></td>
                        <td width="77%" class="txt1"><a href="show.asp?id=<%=rs(0)%>" title="<%=rs(1)%>" class="list6" target=_blank><%=gotTopic(rs(1),40)%></a><%if month(rs(2))=month(date()) and day(rs(2))-day(date())>3 then%><img src="images/news.gif" width="30" height="13" align="absmiddle"><%end if%></td>
                        <td width="19%" class="txt3">[<%=month(rs(2))&"-"&day(rs(2))%>]</td>
                      </tr>
                      <tr>
                        <td background="images/iocn2.gif" height="1" colspan="3"></td>
                      </tr>
                      <%rs.movenext
               			i=i+1
               			loop
               			rs.close%> 
                    </table></td>
                  </tr>
                </table>
                <map name="Map3Map">
                  <area shape="rect" coords="325, 5, 395, 21" href="list.asp?classid=5">
                </map>                <br>
                <table width="400" border="0" align="center" cellpadding="0" cellspacing="3">
                  <tr>
                    <td><img src="images/re.gif" width="400" height="25" border="0" usemap="#Map4"></td>
                  </tr>
                  <tr>
                    <td><table width="100%"  border="0" cellpadding="0" cellspacing="3" class="txt1">
                      <%set rs=conn.execute("select top 5 articleid,title,addtime from article where classid=13 order by addtime desc")
              		i=0
              		do while not rs.eof and i<5%><tr>
                        <td width="4%"><img src="images/iocn1.gif" width="7" height="7"></td>
                        <td width="77%" class="txt1"><a href="show.asp?id=<%=rs(0)%>" title="<%=rs(1)%>" class="list6" target=_blank><%=gotTopic(rs(1),40)%></a><%if month(rs(2))=month(date()) and day(rs(2))-day(date())>3 then%><img src="images/news.gif" width="30" height="13" align="absmiddle"><%end if%></td>
                        <td width="19%" class="txt3">[<%=month(rs(2))&"-"&day(rs(2))%>]</td>
                      </tr>
                      <tr>
                        <td background="images/iocn2.gif" height="1" colspan="3"></td>
                      </tr>
                      <%rs.movenext
               			i=i+1
               			loop
               			rs.close%> 
                    </table></td>
                  </tr>
                </table>
                <br></td>
              <td width="194" align="right" valign="top">
			            
						<table width="171" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/wenxian.gif" width="171" height="28" border="0" usemap="#Map6"></td>
              </tr>
              <tr>
                <td background="images/xr2.gif"><table width="98%"  border="0" cellspacing="2" cellpadding="0">
                  <tr>
                    <td class="txt1"> 
                  <td class="txt1"><%set rs=conn.execute("select top 6 articleid,title from article where classid=6 order by addtime desc")
                  i=0
                  do while not rs.eof and i<6%>
                   ·<a href="show.asp?id=<%=rs(0)%>" title="<%=rs(1)%>" class="list4"><%=rs(1)%></a><br>
      <%rs.movenext
      i=i+1
      loop
      rs.close%> </td>
                  </tr>
                </table></td>
              </tr>
                <tr>
                  <td><img src="images/xr3.gif" width="171" height="9"></td>
                </tr>
          </table><br>
                        <table width="171" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><img src="images/shandong.gif" width="171" height="27" border="0" usemap="#Map11MapMap"></td>
                </tr>
                <tr>
                  <td background="images/xr2.gif"><table width="98%"  border="0" cellspacing="2" cellpadding="0">
                    <tr>
                      <td align="left" class="txt1">
                  <td class="txt1"><%set rs=conn.execute("select top 7 articleid,title from article where classid=15 order by addtime desc")
                  i=0
                  do while not rs.eof and i<7%>
                   ·<a href="show.asp?id=<%=rs(0)%>" title="<%=rs(1)%>" class="list4" target=_blank><%=rs(1)%></a><br>
      <%rs.movenext
      i=i+1
      loop
      rs.close%> </td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td><img src="images/xr3.gif" width="171" height="9"></td>
                </tr>
              </table>
                <map name="Map11MapMap">
                  <area shape="rect" coords="26, 4, 97, 25" href="list.asp?classid=15">
                </map>                <br>
                <table width="171" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><img src="images/xr1.gif" width="171" height="30" border="0" usemap="#Map12"></td>
                </tr>
                <tr>
                  <td align="right" background="images/xr2.gif">
					<table width="98%"  border="0" cellspacing="2" cellpadding="0" id="table1">
                    <tr>
                      <td align="left" class="txt1">
                  <td class="txt1"><%set rs=conn.execute("select top 8 articleid,title from article where classid=12 order by addtime desc")
                  i=0
                  do while not rs.eof and i<8%>
                   ·<a href="show.asp?id=<%=rs(0)%>" title="<%=rs(1)%>" class="list4" target=_blank><%=rs(1)%></a><br>
      <%rs.movenext
      i=i+1
      loop
      rs.close%> </td>
                    </tr>
                  </table>                    </td>
                </tr>
                <tr>
                  <td><img src="images/xr3.gif" width="171" height="9"></td>
                </tr>
              </table>
                <br>
                <table width="171" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td><img src="images/tui.gif" width="171" height="27" border="0" usemap="#Map11"></td>
                  </tr>
                  <tr>
                    <td background="images/xr2.gif"><table width="96%"  border="0" align="right" cellpadding="0" cellspacing="2">
                      <tr>
                        <td align="left" class="txt1"><%set rs=conn.execute("select title,content from article where classid=17")
                        do while not rs.eof
                        %><a href="<%=rs(1)%>" target="_blank" class="list4"><%=rs(0)%> </a><br>
                        <%rs.movenext
                        loop
                        rs.close%></td>
                      </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td><img src="images/xr3.gif" width="171" height="9"></td>
                  </tr>
                </table>                
                <br></td>
            </tr>
          </table>          </td>
        </tr>
      </table>
      <!--#include file="end.inc"--></td>
  </tr>
</table>
<map name="Map2">
  <area shape="rect" coords="517, 2, 587, 22" href="list.asp?classid=3">
</map>
<map name="Map">
  <area shape="rect" coords="325, 3, 394, 21" href="list.asp?classid=9">
</map>
<map name="Map3">
  <area shape="rect" coords="325, 5, 395, 21" href="list.asp?classid=10">
</map>
<map name="Map4">
  <area shape="rect" coords="324, 5, 395, 22" href="list.asp?classid=13">
</map>
<map name="Map6">
  <area shape="rect" coords="38, 6, 111, 24" href="list.asp?classid=6">
</map>
<map name="Map7">
  <area shape="rect" coords="37, 7, 112, 28" href="list.asp?classid=7">
</map>
<map name="Map8">
  <area shape="rect" coords="39, 5, 110, 25" href="list2.asp?classid=8">
</map>
<map name="Map10">
  <area shape="rect" coords="37, 6, 109, 25" href="list2.asp?classid=11">
</map>
<map name="Map11">
  <area shape="rect" coords="16, 3, 87, 24" href="link.asp">
</map>
<map name="Map12">
  <area shape="rect" coords="18, 5, 89, 25" href="list.asp?classid=12">
</map>
</body>
</html>