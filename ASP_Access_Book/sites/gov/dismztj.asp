<!--#include file="rscoon.asp"-->
<% toptitle="�鿴ͳ��" %>
<!--#include file="mztop.asp"-->
<!--#include file="mzfunc.asp"-->
  <%
  dim tjid
  tjid=request.querystring("id")
  if tjid="" or isnull(tjid) then
  response.write"<script>alert('�Բ���,�������ִ���\n�뷵������');history.back();</script>"
    else
	if isnumeric(tjid)<>true or instr(tjid,"or")>0 or instr(tjid,"and")>0 then
	response.write"<script>alert('�Բ���\n�������ܾ�');window.location.href='error.htm';</script>"
          end if
   sql="select tjtitle,tjbody,tjtime from mztj where tjid="&tjid
  set rs=conn.execute(sql)
   if not rs.eof then
response.write"<TABLE width=766 height=100% border=0 align=center cellPadding=0 cellSpacing=1><TBODY><TR>"&_
"<TD width=766 height=404 vAlign=top bgColor=#f5f6fb class=wz7> <TABLE width=100% border=0 align=center cellPadding=0 cellSpacing=1><TBODY><TR>"&_
"<TD height=30 vAlign=top> <p>&nbsp;</p><DIV align='center'><strong>"&rs("tjtitle")&"</strong></DIV>"&_
"<br> <hr align=center width=100% size=1 noshade></TD></TR></TBODY></TABLE><br> &nbsp;&nbsp;&nbsp;"&ubbcode(rs("tjbody"))&"</TD>"&_
"</TR><TR><TD bgColor=#f5f6fb class=wz7 height=20 vAlign=top><div align=right>ͳ������:"&rs("tjtime")&"-----<a href='javascript:window.close()'>���رմ��ڡ�</a></div></TD></TR></TBODY></TABLE>"
 else  response.write"<script>alert('�Բ���,����Ҫ�鿴�Ķ���δ�ҵ�');history.back();</script>" 
end if   end if
  call mzfoot() %>
