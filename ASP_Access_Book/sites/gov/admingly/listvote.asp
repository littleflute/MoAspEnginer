<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
    dim id,stat
	id=request.querystring("id")
	  if id="" or isnull(id) then
	  response.write"<script>alert('对不起,此页面不能刷新');history.back();</script>"
	       else
	sql="select title from stat where id="&id
	   set rs=conn.execute(sql)
	   if not rs.eof then
	   stat=rs("title")
	    end if
		end if
		%>
		
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="30" colspan="2" bgcolor="#B1E1F3"> 
      <div align="center"><strong>【<%=stat%>】的所有选项</strong></div></td>
  </tr>
   <tr> 
    <td width="83%"><div align="center">投票选项</div></td>
    <td width="17%"><div align="center">投票数</div></td>
  </tr>
<%	sql="select votetitle,counts from votes where fromid="&id
	  set rs=conn.execute(sql)
	  if not rs.eof then
	  do while not rs.eof 
%>
 
  <tr> 
    <td><div align="center"><%=rs("votetitle")%></div></td>
    <td><div align="center"><%=rs("counts")%></div></td>
  </tr>
  <% rs.movenext 
  loop 
  else response.write"<tr><td>暂无选项</td></tr>"
  end if%>
<tr>
    <td colspan="2"><div align="right"><a href="javascript:history.go(-1)">【返回】</a></div></td>
  </tr>
</table>
</body>
</html>
