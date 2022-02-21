<% toptitle="民政机构简介..." %>
<!--#include file="mztop.asp"-->
<!--#include file="mzfunc.asp"-->
<!--#include file="topconst.asp"-->
<%
   if frametitle<>"" or framebody<>"" then 
 response.write"<TABLE width='766' height='100%' border=0 align='center' cellPadding=0 cellSpacing=1><TBODY><TR>"&_
"<TD bgColor=#f5f6fb class=wz7 height=404 vAlign=top> <TABLE width='100%' border=0 align='center' cellPadding=0 cellSpacing=1><TBODY><TR><TD height=30 vAlign=top> <p>&nbsp;</p>"&_
"<DIV align='center'><strong>"&frametitle&"</strong> </DIV><br><hr align='center' width='100%' size='1' noshade></TD></TR><TR>"&_
"<TD vAlign=top> <div align='right'>--"&frametime&"发布</div></TD></TR></TBODY></TABLE><br>&nbsp;&nbsp;&nbsp;"&ubbcode(framebody)&"</TD>"&_
"</TR><TR><TD bgColor=#f5f6fb class=wz7 height=20 vAlign=top><div align='right'>-----<a href='javascript:window.close()'>【关闭窗口】</a></div></TD></TR></TBODY></TABLE>"  
 else
response.write"<script>alert('此版块暂无类容');history.back();</script>"
 end if
call mzfoot()
%>

