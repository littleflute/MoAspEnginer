<TABLE  width="100%" border=0 align=left cellPadding=1 cellSpacing=2>
  <TBODY>  <TR>
					<% dim allsort  
					sql="select sort,id from allsort order by id"
					      set rs=conn.execute(sql)
						  allsort=rs.getrows
					%>
                       
      <TD height="188" class=text> <div align="left"> &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,0)%>"> 
          <%=allsort(0,0)%></a><br>
          &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,1)%>"><%=allsort(0,1)%></a><BR>
          &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,2)%>"><%=allsort(0,2)%></a><br>
          &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,3)%>"><%=allsort(0,3)%></a><BR>
          &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,4)%>"><%=allsort(0,4)%></a><br>
          &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,5)%>"><%=allsort(0,5)%></a><BR>
          &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,6)%>"><%=allsort(0,6)%></a><br>
          &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,7)%>"><%=allsort(0,7)%></a><BR>
          &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,8)%>"><%=allsort(0,8)%></a><br>
          &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,9)%>"><%=allsort(0,9)%></a><BR>
          &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,10)%>"><%=allsort(0,10)%></a><br>
          &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,11)%>"><%=allsort(0,11)%></a><BR>
          &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,12)%>"><%=allsort(0,12)%></a><br>
          &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,13)%>"><%=allsort(0,13)%></a><br>
          &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,14)%>"><%=allsort(0,14)%></a><BR>
          &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,15)%>"><%=allsort(0,15)%></a><BR>
          &nbsp;&nbsp; <img src="images/rj-tp1.gif" width="10" height="10" border="0"> 
          <a href="sorts.asp?id=<%=allsort(1,16)%>"><%=allsort(0,16)%></a></p> </div></TD>
                    </TR>
                  </TBODY>
                </TABLE>
