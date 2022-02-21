<% Response.Buffer=True %>
<!--#include file="../inc/company.inc"-->
<% uname=session("cuid")
   set rs=server.createobject("adodb.recordset") 
   if request("del")<>"" then 
   conn.Execute("delete  from cmailbox where id="&request("del"))
   end if
   sql="select * from cmailbox where reid='"&uname&"'and newmail=0"
   rs.open sql,conn,1,1
   mailnum=rs.recordcount
   rs.close                                                                    
   set rs=nothing %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/index.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>天天人才—&gt;站内信箱</title>
</head>
<SCRIPT language=JavaScript src="../inc/window.js"></SCRIPT>
<body topmargin="0" leftmargin="0">
<!--#include file="../inc/top2.htm"--> 
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="358">
    <tr>
      
    <td width="778" height="18" valign="top" colspan="5" bgcolor="#53BEB0"> 　 
    </td>
    </tr>
    <tr>
      
    <td width="139" valign="top" bgcolor="#53BEB0" height="176"> 　 
      <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="118" height="280">
            <tr>
              <td width="100%" height="163" background="../images/stat-bg.GIF" valign="top">
        <p align="center"><br>
        <a href="main.asp">登录首页</a><br>
        <br>
		<a href="register.asp">登录/更新公司资料</a><br>
        <br>
		<a href="publish.asp">发布/更新招聘信息</a><br>
        <br>
        <a href="../changepwd.asp?stype=company" target="_blank">修改登录密码</a><br>
        <br>
        <a href="search.asp">全部人才列表</a>
              </td>
            </tr>
            <tr>
              <td width="100%" height="117" background="../images/stat-bg.GIF" valign="top">
        <p align="center"><br>
        <br>
        <a href="favorite.asp">我的收藏夹<br>
        <br>
        <a href="mailbox.asp">我的信箱</a><br>
        <br>
        <a href="../exit.asp">退出登录</a>
              </td>
            </tr>
          </table>
        </div>
        
      <p align="center">&nbsp; </p>
        <p align="center">　
      </td>
      <td width="20" height="284" valign="top"><img border="0" src="../images/selfk.GIF"></td>
  </center>
      
      <td width="484" height="284" valign="top">
           <% if not isempty(request("page")) then   
	          pagecount=cint(request("page"))   
              else   
	          pagecount=1   
              end if
			  set rs=server.createobject("adodb.recordset") 
			  sql="select * from cmailbox where reid='"&uname&"' order by id desc" 
			  rs.open sql,conn,1,1 
              if rs.eof and rs.bof then   
              response.write"<SCRIPT language=JavaScript>alert('对不起，你的信箱内暂无信件！');"
              response.write"this.location.href='main.asp';</SCRIPT>"      
              end if
              rs.pagesize=10
              if pagecount>rs.pagecount or pagecount<=0 then              
              pagecount=1              
              end if              
              rs.AbsolutePage=pagecount %>   
  
        <div align="left">
         <table border="1" cellpadding="0" cellspacing="0" width="477" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
          <tr>
          <td height="1" colspan="6" valign="bottom" width="473"><font color="#000000">&nbsp;</font></td> 
          </tr>
          <tr>
		 <td width="473" height="6" colspan="6" valign="bottom"><font color="#000000">&nbsp;</font><font color="#000000">
		 <% page_start=(pagecount-1)*rs.pagesize
            if pagecount=1 then page_start=1
		    page_end=rs.pagesize*pagecount
		    if pagecount*rs.pagesize=>rs.recordcount then page_end=rs.recordcount end if%>
         您的信箱内共有[<font color="#ff0000"><%=rs.recordcount%></font>]封站内信件 其中有[<font color="#ff0000"><%=mailnum%></font>]封新邮件                  
         以下是[<font color="red"><%=page_start%>～<%=page_end%></font>]封 页次:[<font color="#0000AE"><%=pagecount%></font>/<%=rs.pagecount%>]</font></td>                
          <tr>
      <td height="3" valign="top" colspan="6" bgcolor="#000000" width="473">
      </td>
          </tr>
		  <tr>
          <td width="42" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">新邮件</p>
          </td>
          <td width="103" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;发信人</td>
          <td width="119" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;标 题</td>         
          <td height="18" bgcolor="#EBEEF3" valign="bottom" width="127"><p align="center">发信时间</p></td>
          <td height="18" bgcolor="#EBEEF3" valign="bottom" width="36">
            <p align="center">回复</td>
          <td height="18" bgcolor="#EBEEF3" valign="bottom" width="36">
            <p align="center">删除</td>
          </tr>
          <% do while not rs.eof %> 
          <tr>
          <td width="42" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center"><% if rs("newmail")="0" then %>
			--><% else %>&nbsp; <% end if %>      
			</td>
          <td width="103" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<a href="javascript:openwin('../person.asp?uid=<%=rs("senduid")%>','top=10,left=300,width=460,height=420')"><% if len(rs("sendname"))>7 then%><%=Left(rs("sendname"),6)%>...<% else%><%=rs("sendname")%><%end if%></a></td>
          <td width="119" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<a href="javascript:openwin('viewmail.asp?id=<%=rs("id")%>','top=10,left=300,width=460,height=420')"><% if len(rs("title"))>7 then%><%=Left(rs("title"),7)%>...<% else%><%=rs("title")%><%end if%></a></td>
          <td width="127" height="18" bgcolor="#EBEEF3" valign="bottom"><p align="center"><%=rs("sdate")%></p></td>
          <td width="36" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center"><a href="javascript:openwin('sendmail.asp?reid=<%=rs("senduid")%>','top=10,left=300,width=460,height=420')">回复</a>
			</td>
          <td width="36" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center"><a href="mailbox.asp?del=<%=rs("id")%>&newmail=<%=rs("newmail")%>">
			<font color="#00006a">删除</font></a></td>
          </tr>
		 <% i=i+1                                                                                                  
          rs.movenext                                                                                                  
          if i>=rs.PageSize then exit do 
		  loop %>
          <tr>
      <td height="3" valign="top" colspan="6" bgcolor="#000000" width="473">
      </td>
          </tr>
          <tr>
          <td width="473" height="8" valign="bottom" colspan="6"><br>
<p align="center"><%if pagecount=1 and rs.pagecount<>pagecount then%><a href="mailbox.asp?page=<%=cstr(pagecount+1)%>">
<font color="#000000">下一页>>></font></a><a>     
<% end if %><% if rs.pagecount>1 and rs.pagecount=pagecount then %></a><a href="mailbox.asp?page=<%=cstr(pagecount-1)%>"> 
<font color="#000000"><<<上一页/font></a><a><%end if%>     
<%if pagecount<>1 and rs.pagecount<>pagecount then%></a><a href="mailbox.asp?page=<%=cstr(pagecount-1)%>"> 
<font color="#000000"><<<上一页/font></a><a>  </a> <a href="mailbox.asp?page=<%=cstr(pagecount+1)%>">    
<font color="#000000">下一页>>></font></a><a>
<% end if 
rs.close                                                                                                
set rs=nothing                                                                                                
conn.close                                                                                                
set conn=nothing %>
            </a>
          </td>
          </tr>
		  </table>
          </div>
</td>
  <center>
      <td width="1" height="284" valign="top" bgcolor="#00006A"></td>
      <td width="126" height="284" valign="top" bgcolor="#F3F3F3">　</td>
    </tr>
    <tr>
      
    <td height="1" valign="top" colspan="5" bgcolor="#53BEB0" width="778"> 
      <p align="center"> </td>
    </tr>
  </center>
    </table>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="1">
    <tr>
      <td width="743" height="3" valign="top">
      <p align="center">
      </td>
    </tr>
    <tr>
      <td width="743" height="1" valign="top">
      <p align="center"><script language="javascript" src="../inc/copyright.js"></script>
      </td>
    </tr>
    <tr>
      <td width="743" height="3" valign="top">
      </td>
    </tr>
  </table>
</body>
</html>
 

