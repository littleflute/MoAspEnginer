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
              <TD align=middle colSpan=2><FONT color=#333333 
            face=����><B>��ӭʹ�������ռ��䣬</B></FONT></TD>
            </TR>
            <TR> 
              <TD align=middle bgColor=#eeeeee colSpan=2 vAlign=center><a href="duan.asp"><IMG 
            alt=�ռ��� border=0 src="duanx.files/inboxpm.gif"></a> &nbsp; <IMG 
            alt=������ border=0 src="duanx.files/outboxpm.gif"> &nbsp; <A 
            href="duanadd.asp"><IMG 
            alt=������Ϣ border=0 src="duanx.files/newpm.gif"></A></TD>
            </TR>
            <%
uer=request.QueryString("id")
sql_news="Select * From duanx where id="&uer
set rs_news=conn.execute(sql_news)
if not rs_news.eof then
	 %>
            <TR> 
              <TD colspan="2" align=middle vAlign=center bgColor=#eeeeee>��<%=rs_news("time")%>���������͵���Ϣ�� </TD>
            </TR>
            <TR> 
              <TD width="9%" align=middle vAlign=center bgColor=#eeeeee><font color="#333333" face="����">����</font></TD>
              <TD width="91%" align=middle vAlign=center bgColor=#eeeeee><font 
            color=#333333 face=����><%= rs_news("title") %></font></TD>
            </TR>
            <TR> 
              <TD align=left vAlign=top bgColor=#eeeeee><FONT 
            color=#333333 face=����> ����</FONT></TD>
              <TD align=middle vAlign=center bgColor=#eeeeee><font 
            color=#333333 face=����><%= rs_news("content") %></font></TD>
            </TR>
            <%	
	end if
rs_news.close
%>
            <TR bgColor=#ccff99> 
              <TD align=middle colSpan=2 vAlign=center><FONT color=#333333 
            face=����>ɾ����������Ϣ</FONT></TD>
            </TR>
          </TBODY>
        </TABLE>
    </TD></TR></TBODY></TABLE>
<br>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width="95%">
  <TBODY>
    <TR> 
      <TD> </TD>
  </TBODY>
</TABLE>
</BODY></HTML>
 <%   
id=request.QueryString("id")   
Set rs = Server.CreateObject("ADODB.Recordset")          
rs.Open "Select * From duanx where id="&id, conn, 3, 2  
rs("flag")="��"
rs.update
rs.close %>  