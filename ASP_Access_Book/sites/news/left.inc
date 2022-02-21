<table width="171" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td colspan="3" class="txt2"><img src="images/sou.jpg" width="171" height="35"></td>
            </tr>
            <form action=search.asp method=get target=_blank><tr>
              <td width="8" align="right">¡¡</td>
              <td width="136" height="30">
                <input name="key" type="text" class="f1" size="18">
              </td>
              <td width="27"><input type=image src="images/go.gif" width="20" height="20" border="0"></td>
            </tr></form>
          </table> 
            <table width="171" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><img src="images/ling.gif" width="171" height="28" border="0" usemap="#Map7"></td>
            </tr>
            <tr>
              <td background="images/bg1.jpg"><table width="98%"  border="0" cellspacing="2" cellpadding="0">
                <tr>
                  <td class="txt1">
                  <td class="txt1"><%set rs=conn.execute("select top 4 articleid,title from article where classid=19 order by addtime desc")
                  i=0
                  do while not rs.eof and i<4%>
                   ¡¤<a href="show.asp?id=<%=rs(0)%>" title="<%=rs(1)%>" class="list5"><%=rs(1)%></a><br>
      <%rs.movenext
      i=i+1
      loop
      rs.close%> </td>
                </tr>
              </table></td>
            </tr>
          </table>
          
          <table width="171" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><img src="images/guo.gif" width="171" height="32" border="0" usemap="#Map7Map"></td>
            </tr>
            <tr>
              <td background="images/bg2.jpg"><table width="98%"  border="0" cellspacing="2" cellpadding="0">
                <tr>
                  <td class="txt1"><%set rs=conn.execute("select top 7 articleid,title from article,class where class.parentid=4 and class.classid=article.classid order by addtime desc")
                  i=0
                  do while not rs.eof and i<7%>
                   ¡¤<a href="show.asp?id=<%=rs(0)%>" title="<%=rs(1)%>" class="list5"><%=rs(1)%></a><br>
      <%rs.movenext
      i=i+1
      loop
      rs.close%> </td>
                </tr>
              </table></td>
            </tr>
          </table>
          <map name="Map7Map">
            <area shape="rect" coords="37, 7, 112, 28" href="list.asp?classid=4">
          </map> <br> 