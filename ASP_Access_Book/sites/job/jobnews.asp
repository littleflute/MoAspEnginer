<% Response.Buffer=True %> 
<!--#include file="inc/dbconn.inc"-->
<% set rs=server.createobject("adodb.recordset")
sql="select * from jobnews order by id desc" 
rs.open sql,conn,1,1 
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1   
end if
rs.pagesize=15 
rs.AbsolutePage=pagecount%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/index.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>新闻资讯|NEWS</title>
</head>
<body>

<div align="center">
  <table border="1" cellpadding="0" cellspacing="0" width="430" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td width="270" height="17" valign="bottom" background="images/t-bg1.gif">
        <p align="left">&nbsp;新闻资讯|NEWS</td>
      <td width="154" height="17" valign="bottom" background="images/t-bg1.gif">
        <p align="center">共<font color="#0000AE"><%=rs.recordcount%></font>条新闻 页次:<font color="#0000AE"><%=pagecount%></font>/<%=rs.pagecount%></td>   
    </tr>
  <center>
      <tr bgcolor="#fffff4"> 
        <td width="448" valign="bottom" colspan="2"> 
          <% do while not rs.eof %>
          &nbsp;==>&nbsp;<a href="viewnews.asp?id=<%=rs("id")%>" target="_blank"> 
          <% if len(rs("title"))>18 then%>
          <%=left(rs("title"),15)%>... 
          <% else%>
          <%=rs("title")%> 
          <%end if%>
          </a><I> [<%=rs("idate")%>]</I>&nbsp;点击<font color="#000091"><%=rs("click")%></font>次<BR>
          <% n=n+1                                                                     
           rs.movenext                                                                     
          if n>=rs.pagesize then exit do                                                                     
           loop %>
        </td>
    </tr>
  </center>
    <tr>
<td width="448" height="17" valign="bottom" background="images/t-bg1.gif" colspan="2">  
<p align="center"><%if pagecount=1 and rs.pagecount<>pagecount then%><a href="jobnews.asp?page=<%=cstr(pagecount+1)%>">
<font color="#000000">下一页</font><a>
<% end if %><% if rs.pagecount>1 and rs.pagecount=pagecount then %><a href="jobnews.asp?page=<%=cstr(pagecount-1)%>">
<font color="#000000">上一页</font><a><%end if%>
<%if pagecount<>1 and rs.pagecount<>pagecount then%><a href="jobnews.asp?page=<%=cstr(pagecount-1)%>">
<font color="#000000">上一页</font><a> <a href="jobnews.asp?page=<%=cstr(pagecount+1)%>">
<font color="#000000">下一页</font><a>
<%end if%>
	</td>
	</tr>
  </table>
</div>

<p align="center">【<a href="javascript:window.close()">关闭窗口</a>】</p>

</body>

</html>
