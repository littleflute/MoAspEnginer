<%@ CODEPAGE = "936" %>
<!--#include file="conn.asp"-->
<%
  dim rs
  set rs = server.createobject("adodb.recordset")
%>
<html>
<head>
<title>最新软件列表</title>
<META content="text/html; charset=gb2312" http-equiv="Content-Type">
<LINK rel="stylesheet" href="images/style1.css">
</head>
<!--#include file="head.asp"-->
<div align="center"> 
  <center>
    <!--#include file="lanmu.asp"-->
    <table border="0" cellpadding="0" cellspacing="0"  width="770" height="1" id="AutoNumber1">
      <tr> 
        <td width="100%" height="1" vAlign="top" bgcolor="#ffffff"> 
          <div align="center"> <b></b>最新软件更新<b>－－分类最新更新列表</b> 
            <select onchange="javascript:window.open(this.options[this.selectedIndex].value,&quot;_top&quot;)" name="select">
              <option selected>全部分类</option>
              <%      
	sql="select class,classid from class"      
	rs.open sql,conn,1,1      
	if rs.bof and rs.eof then      
%>
              <OPTION value>Not Record!</OPTION>
              <%else%>
              <%      
	do while not rs.eof      
%>
              <OPTION <%if request("classid")<>"" then%><%if cint(rs("classid"))=cint(request("classid")) then%> <%end if%><%end if%>value="class.asp?classid=<%=rs("classid")%>"><%=rs("class")%></OPTION>
              <%      
	rs.movenext      
	loop      
	end if      
	rs.close      
%>
            </SELECT>
            <br>
            <%
sql="select count(id) from download where Year(dateandtime)=Year(now()) and Month(dateandtime)=Month(now()) and Day(dateandtime)=Day(now())"
today=conn.execute(sql)(0)
if today=0 then
response.write "今日更新软件数量：暂无<br>"
else
response.write "今日更新软件数量：<font color=#ff0000><b>"&today&"</b></font> 个"
end if
%>
            <br>
            <table cellSpacing="0" cellPadding="3" width="770" border="1" bordercolordark=#ffffff  bordercolor=#84b6ad  height="10">
              <tr align="center" background="pic/bar.gif"> 
                <td width="100" height="4"> 
                  <p align="center"> 类别</p>
                </td>
                <td width="388" height="4">软件名</td>
                <td width="150" height="4">更新时间</td>
                <td width="60" height="4">总下载次数</td>
              </tr>
            </table>
          </div>
      </tr>
    </table>
    <table width="770"  height="228" border="1" cellPadding="3" cellSpacing="0"  bordercolor=#84b6ad bordercolordark=#ffffff bgcolor="#FFFFFF">
      <%    
	sql="select top 50 "    
	sql=sql&"download.id,showname,bb,dateandtime,hits,download.classid,download.Nclassid,Nclass.Nclass "    
	sql=sql&" from download,Nclass where download.Nclassid=Nclass.Nclassid "    
	sql=sql&" order by download.dateandtime desc"    
	rs.open sql,conn,1,1    
	if rs.eof and rs.bof then    
%>
      <tr> 
        <td width="108">没有任何软件</td>
      </tr>
      <%else%>
      <%do while not rs.eof%>
      <tr align="center" onmouseover="this.bgColor='#8D8C72';this.style.cursor='hand';" onmouseout="this.bgColor='#ffffff';"> 
        <td width="108" height="16"> <p align="center"> <img src="pic/lits.gif" width="13" height="14">  <a href='class.asp?classid=<%=rs("classid")%>&Nclassid=<%=rs("Nclassid")%>'> 
            <%=rs("Nclass")%></a></p>
        </td>
        <td width="419" height="16" align="left"> <a href="software.asp?id=<%=rs("id")%>"><%=rs("showname")%></a></td>
        <td width="158" height="16">( 
          <%if (month(rs("dateandtime"))&day(rs("dateandtime")))=(month(date())&day(date())) then%>
          <font color=red> 
          <%else%>
          <font color='#999999'> 
          <%end if%>
          <%=month(rs("dateandtime"))%>月<%=day(rs("dateandtime"))%>日</font>)</font></td>
        <td width="68" height="16"><font color=green><%=rs("hits")*3%></font></td>
        <%      
	rs.movenext          
	loop                     
	end if                    
	rs.close                     
%>
      </tr>
    </table>
    <br>
  </center>
</div>
<table cellSpacing="0" cellPadding="0" width="770" bgColor="#000000" border="0" align="center">
  <tr> 
    <td><img height="1" src="" width="770"></td>
  </tr>
</table>
<!--#include file="foot.asp"-->
</body>
</html>