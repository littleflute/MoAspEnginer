<!--#include file="conn.asp"-->
<%
homepath = "kejian/" '��������Ŀ¼
rows = 10 '����Ҫ��ʾ������
showlength = 10 '�������ϱ�����ʾ���������������ֽ�����"..."����

set rs = server.createobject("adodb.recordset")
sql = "select * from main,teacher where main.idofteacher=teacher.teacherid order by dateandtime desc"
rs.open sql,conn,1,1
for i = 1 to rows
 if rs.eof then exit for
 dateandtime = rs("dateandtime")
 updatedate = year(dateandtime)&"/"&month(dateandtime)&"/"&day(dateandtime)
 if len(rs("title")) > showlength then
  filetitle = left(rs("title"),showlength-2)&"..."
 else
  filetitle = rs("title")
 end if
%>
document.write("<a href=# onclick=javascript:showdetail(<%=rs("mainid")%>); title='���⣺<%=rs("title")%>\n��ʦ��<%=rs("teacher")%>  �������ڣ�<%=updatedate%>'><%=filetitle%></a><br>");
<%
 rs.movenext
next
rs.close
set rs = nothing
conn.close
set conn = nothing
%>
function showdetail(id)
{
 window.open("<%=homepath%>detail.asp?id="+id,"detail","width=400,height=400,resizable=no,scrollbars=yes,menubar=no,status=no");
}