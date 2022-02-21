<%@ CODEPAGE = "936" %>
<!--#include file="conn.asp"-->

<%
dim shu,news,classid,changdu
changdu=24 '显示长度
if request.querystring("changdu")="" or request.querystring("changdu")=0 then
changdu=24
else
changdu=CINT(request.querystring("changdu"))
end if

if request.querystring("shu")="" or request.querystring("shu")=0 then
shu=8
else
shu=CINT(request.querystring("shu"))
end if


set rs=server.createobject("adodb.recordset")

if request("classid")<>"" then 
classid = request("classid")
set rs=conn.execute("SELECT Nclass.Nclassid,Nclass.Nclass,download.id,download.showname,download.classid,download.Nclassid,download.dateandtime,download.hots FROM download,Nclass where download.classid="&cstr(classid)&" and download.Nclassid=Nclass.Nclassid order by download.id desc")
else 
set rs=conn.execute("SELECT Nclass.Nclassid,Nclass.Nclass,download.id,download.showname,download.classid,download.Nclassid,download.dateandtime,download.hots FROM download,Nclass where download.Nclassid=Nclass.Nclassid order by download.id desc")
end if
if rs.eof and rs.bof then %>
     document.write('<p align="center">Sorry! 没有软件</p>');
 
 <% else  
	news=1
do while not rs.eof%>
document.write('<table width="100%" border="0" cellspacing="0" cellpadding="0"> <TR><TD><img src="pic/lits.gif">  <A href="wangruan/software.asp?id=<%=rs("id")%>"><%if len(rs("showname"))>changdu then%><%=left(rs("showname"),changdu)%>...<%else%><%=rs("showname")%><%end if%></A></TD></TR></table>');						
<%    news=news+1
     if news>shu then exit do 
      rs.movenext 
     	loop  

end if
	Rs.Close
set Rs=nothing
%>

