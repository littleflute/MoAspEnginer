<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn.asp"-->

<HTML><HEAD><TITLE>����Ϣ</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="duanx.files/forum.css" rel=stylesheet>
<META content="MSHTML 5.00.3315.2870" name=GENERATOR></HEAD>
<BODY aLink=#333333 bgColor=#eeeeee link=#333333 
onkeydown="if(event.keyCode==13 &amp;&amp; event.ctrlKey)messager.submit()" 
topMargin=10 vLink=#333333>

<TABLE align=center bgColor=#009933 border=0 cellPadding=0 cellSpacing=0 
width="95%">
  <TBODY>
  <TR>
    <TD>
      <TABLE border=0 cellPadding=3 cellSpacing=1 width="100%">
          <TBODY>
            <TR bgColor=#ccff99> 
              <TD align=middle colSpan=4><FONT color=#333333 
            face=����><B>��ӭʹ�������ռ��䣬</B></FONT></TD>
            </TR>
            <TR> 
              <TD align=middle bgColor=#eeeeee colSpan=4 vAlign=center><a href="duan.asp"><IMG 
            alt=�ռ��� border=0 src="duanx.files/inboxpm.gif"></a> &nbsp; <A 
            href="duanadd.asp"><IMG 
            alt=������Ϣ border=0 src="duanx.files/newpm.gif"></A></TD>
            </TR>
            <TR> 
              <TD align=middle bgColor=#eeeeee vAlign=center width="18%"><FONT color=#333333 
            face=����><B>����</B></FONT></TD>
              <TD align=middle bgColor=#eeeeee vAlign=center><FONT color=#333333 
            face=����><B>����</B></FONT></TD>
              <TD align=middle bgColor=#eeeeee vAlign=center width="9%"><FONT color=#333333 
            face=����><B>�Ѷ�</B></FONT></TD>
              <TD align=middle bgColor=#eeeeee vAlign=center width="25%">ʱ��</TD>
            </TR>
    <%
uer=session("AdminUID")
sql_news="Select * From duanx where [to]="&"'"&uer&"'order by flag"
set rs_news=conn.execute(sql_news)
if not rs_news.eof then
	rs_news.movefirst
	do while not rs_news.eof %>
            <TR> 
              <TD align=middle bgColor=#eeeeee width="15%" vAlign=center><FONT 
            color=#333333 face=����>&nbsp; <%= rs_news("from") %></FONT></TD>
              <TD align=middle bgColor=#eeeeee width="55%" vAlign=center><a href=duanx.asp?id=<%= rs_news("id")%>><%= rs_news("title") %></a></TD>
              <TD width="9%" align=middle vAlign=center bgColor=#eeeeee>
                <% If rs_news("flag")="��" Then %>
                <font color="#FF0000"><%=rs_news("flag")%></font><% Else %><%=rs_news("flag")%><% End If %>

</TD>
              <TD width="15%" align=middle vAlign=center bgColor=#eeeeee><%=rs_news("time")%></TD>
            </TR>
			<%	rs_news.movenext
	loop
	end if
rs_news.close
%>
            <TR bgColor=#ccff99> 
              <TD align=middle colSpan=4 vAlign=center><FONT color=#333333 
            face=����><a href="duandel.asp?id=<%=uer%>">ɾ�����еĶ���Ϣ</a></FONT></TD>
            </TR>
          </TBODY>
        </TABLE>
    </TD></TR></TBODY></TABLE>
<br>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width="95%">
  <TBODY>
    <TR> 
      <TD> <P align=center><A 
    href="http://swrh.whu.edu.cn/xgol/forum/</td"></A>Copyright&copy; 2003  
          all rights reserved</P></TD>
  </TBODY>
</TABLE>
</BODY></HTML>
