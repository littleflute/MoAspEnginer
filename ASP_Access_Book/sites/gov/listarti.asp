<!--#include file="rscoon.asp"-->
<% toptitle="�鿴��Ϣ" %>
<!--#include file="mztop.asp"-->
<!--#include file="mzfunc.asp"-->
<TABLE width="766" height="100%" border=0 align="center" cellPadding=0 cellSpacing=1>
  <TBODY>
    <TR> 
      <TD bgColor=#f5f6fb class=wz7 height=404 vAlign=top> 
        <TABLE width="766" border=0 align="center" cellPadding=0 cellSpacing=1>
          <TBODY>
            <TR> 
			<%
			     id=request.querystring("id")
				 if id="" or isnull(id) or isnumeric(id)<>true then
				 response.write"<script>alert('��������\n�뷵������');history.back();</script>"
			          else
			   sql="select title, body,times from allarti where id="&id 
			         set rs=conn.execute(sql)
					 end if
					 if rs.eof then
					 response.write"<script>alert('��������\n�뷵������');history.back();</script>"
			             else
			%>
              <TD height=53 vAlign=top> <br>
                <DIV align=center> 
                  <font size="4"><%=rs("title")%></font>
                </DIV></TD>
            </TR>
          </TBODY>
        </TABLE>
       &nbsp;&nbsp;&nbsp;<%=UBBCode(rs("body"))%></TD>
    </TR>
    <TR>
      <TD bgColor=#f5f6fb  height=20 vAlign=top><div align="right"><a href="javascript:window.close()">���رմ��ڡ�</a></div></TD>
    </TR>
  </TBODY>
</TABLE>
<% end if 
  call mzfoot() %>
