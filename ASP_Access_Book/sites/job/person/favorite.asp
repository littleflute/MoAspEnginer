<% Response.Buffer=True %>
<!--#include file="../inc/person.inc"-->
<% uname=session("puid")
   if not isempty(request("page")) then   
	pagecount=cint(request("page"))   
    else   
	pagecount=1   
    end if
   set rs=server.createobject("adodb.recordset")
   if request("del")<>"" then 
   conn.Execute("delete  from pfavorite where uname='"&uname&"' and fuid='"&request("del")&"'")
   end if
   sql="select * from pfavorite where uname='"&uname&"' order by id desc"
   rs.open sql,conn,1,1
   if rs.eof and rs.bof then   
   response.write"<SCRIPT language=JavaScript>alert('对不起，你的收藏夹为空！');"
   response.write"javascript:history.go(-1);</SCRIPT>"
   end if
   fnum=rs.recordcount
   rs.pagesize=10
   fpnum=rs.pagecount
   if pagecount>rs.pagecount or pagecount<=0 then              
   pagecount=1              
   end if              
   rs.AbsolutePage=pagecount
   page_start=(pagecount-1)*rs.pagesize
   if pagecount=1 then page_start=1
   page_end=rs.pagesize*pagecount
   if pagecount*rs.pagesize=>rs.recordcount then page_end=rs.recordcount end if
   do while not rs.eof  
   id=rs("id")
   if fuid="" then
      fuid=rs("fuid")
   else
      fuid=fuid & "," & rs("fuid")
   end if
   i=i+1
   rs.movenext
   if i>=rs.PageSize then exit do
   loop
   rs.close                                                                    
   set rs=nothing
   fuid=split(fuid,",")
   set rs=server.createobject("adodb.recordset") %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/index.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>天天人才—&gt;我的收藏夹</title>
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
        <a href="register.asp">登录求职简历</a><br>
        <br>
        <a href="modify.asp">更新求职简历</a><br>
        <br>
        <a href="../changepwd.asp?stype=person" target="_blank">修改登录密码</a><br>
        <br><a href="search.asp">全部职位列表</a>
              </td>
            </tr>
            <tr>
              <td width="100%" height="117" background="../images/stat-bg.GIF" valign="top">
        <p align="center"><br>
        <br><a href="favorite.asp">我的收藏夹</a><br>
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
        <div align="left">
         <table border="1" cellpadding="0" cellspacing="0" width="477" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
          <tr>
          <td height="1" colspan="5" valign="bottom" width="473"><font color="#000000">&nbsp;</font></td> 
          </tr>
          <tr>
		 <td width="473" height="6" colspan="5" valign="bottom"><font color="#000000">&nbsp;</font><font color="#000000">
		 您的收藏夹内共有[<font color="#ff0000"><%=fnum%></font>]条记录</font>&nbsp;以下是[<font color="red"><%=page_start%>～<%=page_end%></font>]条 页次:[<font color="#0000AE"><%=pagecount%></font>/<%=fpnum%>]</td><tr>   
      <td height="3" valign="top" colspan="5" bgcolor="#000000" width="473">
      </td>
          </tr>
          <tr>
          <td width="144" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;公司名称</td>
          <td width="162" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;招聘职位</td>
          <td height="18" bgcolor="#EBEEF3" valign="bottom" width="78"><p align="center">发布日期</p></td>
          <td height="18" bgcolor="#EBEEF3" valign="bottom" width="37">
            <p align="center">发信</td>
          <td height="18" bgcolor="#EBEEF3" valign="bottom" width="32">
            <p align="center">删除</td>
          </tr>
           <% for i=0 to ubound(fuid)
		   sql="select * from company where uname='"&(fuid(i))&"'" 
		   rs.open sql,conn,1,1 %>
          <tr>
          <td width="144" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<a href="javascript:openwin('../company.asp?uid=<%=rs("uname")%>','top=10,left=300,width=460,height=420')"><% if len(rs("cname"))>10 then%><%=Left(rs("cname"),10)%>...<%else%><%=rs("cname")%><%end if%></a></td>
          <td width="162" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<a 
    href="javascript:openwin('../job.asp?uid=<%=rs("uname")%>','top=10,left=300,width=460,height=420')"><%=rs("job")%>
   </a></td>
          <td width="78" height="18" bgcolor="#EBEEF3" valign="bottom"><p align="center">[<%=rs("idate")%>]</p></td>
          <td width="37" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center"><a href="javascript:openwin('sendmail.asp?reid=<%=rs("uname")%>','top=20,left=200,width=450,height=400')"><font face="Wingdings" color="#000060" size="2">*</font></a></td>
          <td width="32" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center"><a href="favorite.asp?del=<%=rs("uname")%>">
            <font color="#000060">删除</font></a></td>
          </tr>
          <% rs.close                                                                    
             next %>
          <tr>
      <td height="3" valign="top" colspan="5" bgcolor="#000000" width="473">
      </td>
          </tr>
          <tr>
          <td width="473" height="8" valign="bottom" colspan="5">
       
		  <p align="center"><br>
		  <%if pagecount=1 and fpnum<>pagecount then%><a href="favorite.asp?page=<%=cstr(pagecount+1)%>">
          <font color="#000000">下一页>>></font></a><a>     
        <% end if %><% if fpnum>1 and fpnum=pagecount then %></a><a href="favorite.asp?page=<%=cstr(pagecount-1)%>"> 
        <font color="#000000"><<<上一页/font></a><a><%end if%>    
        <%if pagecount<>1 and fpnum<>pagecount then%></a><a href="favorite.asp?page=<%=cstr(pagecount-1)%>"> 
        <font color="#000000"><<<上一页/font></a><a>  </a> <a href="favorite.asp?page=<%=cstr(pagecount+1)%>">   
        <font color="#000000">下一页>>></font></a><a>
        <% end if                                                                                               
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
      <td width="126" height="284" valign="top" bgcolor="#F3F3F3"></td>
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
 

