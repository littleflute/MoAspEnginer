<!--#include file="rscoon.asp"-->
<% toptitle="�������..." %>
<!--#include file="mztop.asp"-->
<!--#include file="mzfunc.asp"-->
  <%
	 dim id
	 id=request.querystring("id")
	 if id="" or isnull(id) or isnumeric(id)<>true then
	 response.write"<script>alert('�Բ���,���Ĳ������ִ���');history.back();</script>"
            end if
	sql="select nstitle,nsbody,writers,times from news where id="&id
	set rs=conn.execute(sql)
    if not rs.eof then
%>
  <TABLE width="766" height="100%" border=0 align="center" cellPadding=0 cellSpacing=1>
    <TBODY>
    <TR> 
      <TD bgColor=#f5f6fb class=wz7 height=404 vAlign=top> <TABLE width="100%" border=0 align="center" cellPadding=0 cellSpacing=1>
            <TBODY>
            <TR> 
              <TD height=30 vAlign=top> <p>&nbsp;</p>
                <DIV align="center"><strong> <%=rs("nstitle")%></strong> </DIV><br>
                  <hr align="center" width="100%" size="1" noshade></TD>
            </TR>
            <TR> 
              <TD vAlign=top> <div align="right">--<%=rs("writers")%>����</div></TD>
            </TR>
          </TBODY>
        </TABLE><br>
         &nbsp;&nbsp;&nbsp;
		 <%=ubbcode(rs("nsbody"))%></TD>
    </TR>
    <TR> 
      <TD bgColor=#f5f6fb class=wz7 height=20 vAlign=top><div align="right">��������:<%=rs("times")%>-----<a href="javascript:window.close()">���رմ��ڡ�</a></div></TD>
    </TR>
  </TBODY>
</TABLE>
<%  else
response.write"<script>alert('�Բ���,����Ҫ�鿴�Ķ��󲢲�����\n�벻Ҫ����ĵص�ַ����������');history.back();</script>"
 end if
call mzfoot()
%>

