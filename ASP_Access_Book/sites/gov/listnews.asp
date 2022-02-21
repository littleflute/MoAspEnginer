<!--#include file="rscoon.asp"-->
<% toptitle="新闻浏览..." %>
<!--#include file="mztop.asp"-->
<!--#include file="mzfunc.asp"-->
  <%
	 dim id
	 id=request.querystring("id")
	 if id="" or isnull(id) or isnumeric(id)<>true then
	 response.write"<script>alert('对不起,您的操作出现错误');history.back();</script>"
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
              <TD vAlign=top> <div align="right">--<%=rs("writers")%>报导</div></TD>
            </TR>
          </TBODY>
        </TABLE><br>
         &nbsp;&nbsp;&nbsp;
		 <%=ubbcode(rs("nsbody"))%></TD>
    </TR>
    <TR> 
      <TD bgColor=#f5f6fb class=wz7 height=20 vAlign=top><div align="right">发表日期:<%=rs("times")%>-----<a href="javascript:window.close()">【关闭窗口】</a></div></TD>
    </TR>
  </TBODY>
</TABLE>
<%  else
response.write"<script>alert('对不起,您所要查看的对象并不存在\n请不要随意的地地址栏输入数字');history.back();</script>"
 end if
call mzfoot()
%>

