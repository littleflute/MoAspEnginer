<%@ CODEPAGE = "936" %>
<!--#include file="conn.asp"-->
<%
  dim rs
  set rs = server.createobject("adodb.recordset")
%>
<html>
<head>
<title>软件排行榜</title>
<META content="text/html; charset=gb2312" http-equiv="Content-Type">
</head><!--#include file="head.asp"-->
<!--#include file="lanmu.asp"-->
<br>
<table border="0" cellpadding="0" cellspacing="0"  width="770" height="19" id="AutoNumber1" bgcolor="#FFFFFF">
  <tr bgcolor="#FFFFFF"> 
      <td width="774" height="40" valign=top colspan="3"> <br>
        
      <div align=center><b>排行榜总览</b>(实时数据)－－－<font color="#FF0000">分类下载排行</font><font color="#000000" size="3"><b>--&gt; 
        </b></font> 
        <select onchange="javascript:window.open(this.options[this.selectedIndex].value,&quot;_top&quot;)" name="select">
             <option selected>全部分类</option>
             <%      
	sql="select class,classid from class"      
	rs.open sql,conn,1,1      
	if rs.bof and rs.eof then      
%><OPTION value>Not Record!</OPTION><%else%> <%      
	do while not rs.eof      
%><OPTION <%if request("classid")<>"" then%><%if cint(rs("classid"))=cint(request("classid")) then%> <%end if%><%end if%>value="class.asp?classid=<%=rs("classid")%>"><%=rs("class")%></OPTION><%      
	rs.movenext      
	loop      
	end if      
	rs.close      
%></SELECT> </div></td>
</tr>
<tr>
<td width="774" height="1" bgcolor="#FFFFFF" valign=top colspan="3">
</td>
</tr>
<tr>
      <td width="258" height="148" bgcolor="#FFFFFF" valign=top> 
        <TABLE border=1 bordercolordark=#ffffff  bordercolor=#84b6ad cellpadding=0 cellspacing=0 width=235 align="center">
        
<TR align=middle>
            
          <TD valign="top" width="234"> 
            <div align="center">
<center>
                  <TABLE cellpadding=4 cellspacing=0 width=232 height="100"  bgcolor="#FFFFFF">
                    <tr>
                      
                    <td width="231" height="9" align="center" background="pic/bar.gif" bgColor="#FFCC00"> 
                      <b>本日下载排行(TOP50)</b> </td>
</tr>
<tr>
<td align="left" width="231"><%
	dim tdate
	tdate=year(Now()) & "-" & month(Now()) & "-" & day(Now())
	sql="select top 50 id,showname,bb,dayhits from download where "
	sql=sql&" lasthits="&tdate&" and stop=0 and dayhits>0 "
	sql=sql&" order by dayhits desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
response.write ("本日没有下载")
else
do while not rs.eof
response.write ("<li><A href=software.asp?id="&rs("id")&" title=今日下载："&rs("dayhits")&"次>"&rs("showname")&"</A></li>")
i=i+1
if i>=50 then exit do
rs.movenext
loop
i="0"
end if
rs.close
%>　</td>
</tr>
</table>
</center>
</div>
</TABLE>
    </td>
      <td width="258" height="148" valign=top> 
        <TABLE border=1 bordercolordark=#ffffff  bordercolor=#84b6ad cellpadding=0 cellspacing=0 width=235 align="center">
        <TR align=middle> 
          <TD valign="top" width="234"> 
                <TABLE cellpadding=4 cellspacing=0 width=232 height="100"  bgcolor="#FFFFFF">
                  <tr> 
                    <td width="231" height="15" align="center" background="pic/bar.gif"> 
                      <b>本周下载排行(TOP50)</b> </td>
                  </tr>
                  <tr> 
                    <td align="left" width="231">
                      <%           
	OldWeek = WeekDay(Date())-1           
	If OldWeek = 0 Then OldWeek = 7           
	OldWeek = Date()-OldWeek           
	NewWeek = Date()+(9-WeekDay(Date()))           
	sql="select top 50 id,showname,bb,weekhits from download where "           
	sql=sql&"  (lasthits < " & NewWeek & ") And (lasthits > " & OldWeek & ") and stop=0 and weekhits>0 "           
	sql=sql&" order by weekhits desc"           
	rs.open sql,conn,1,1           
	if rs.eof and rs.bof then           
response.write ("本周没有下载")
else
do while not rs.eof
response.write ("<li><A href=software.asp?id="&rs("id")&" title=本周下载："&rs("weekhits")&"次>"&rs("showname")&"</A></li>")  
i=i+1                        
if i>=50 then exit do                        
rs.movenext                        
loop                        
i="0"                        
end if
rs.close
%>
                      　</td>
                  </tr>
                </table>
              </center>
            </div></TABLE>
    </td>
      <td width="258" height="148" valign=top> 
        <TABLE border=1 bordercolordark=#ffffff  bordercolor=#84b6ad cellpadding=0 cellspacing=0 width=235 align="center">
        <TR align=middle> 
          <TD valign="top" width="234"> <div align="center"> 
              <center>
                <TABLE cellpadding=4 cellspacing=0 width=232 height="100"  bgcolor="#FFFFFF">
                  <tr> 
                    <td width="231" height="13" align="center" background="pic/bar.gif"> 
                      <b>累计下载排行(TOP50)</b> </td>
                  </tr>
                  <tr> 
                    <td align="left" width="231"> 
                      <%                   
	sql="select top 50 id,showname,bb,hits,dayhits from download "                   
	sql=sql&" order by hits desc"                   
	rs.open sql,conn,1,1                   
	if rs.eof and rs.bof then
response.write ("没有下载")
else
do while not rs.eof
response.write ("<li><A href=software.asp?id="&rs("id")&" title=总下载："&rs("hits")*3&"次-本日下载："&rs("dayhits")&"次>"&rs("showname")&"</A></li>")
	rs.movenext                                     
	loop                                     
	end if                                     
	rs.close                                     
%>
                      　</td>
                  </tr>
                </table>
              </center>
            </div></TABLE>
        
    </td>
</tr>
</table>
<br>
<table cellSpacing="0" cellPadding="0" width="770" bgColor="#000000" border="0" align="center">
<tr>
<td><img height="1" src="" width="770"></td>
</tr>
</table><!--#include file="foot.asp"-->
</body>
</html>