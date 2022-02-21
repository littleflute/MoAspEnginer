<% Response.Buffer=True %>
<!--#include file="../inc/person.inc"-->
<% uname=session("puid")
set rs=server.createobject("adodb.recordset")
sql="select * from pmailbox where reid='"&uname&"'and newmail=0"
rs.open sql,conn,1,1
mailnum=rs.recordcount
rs.close                                                                    
set rs=nothing
set rs=server.createobject("adodb.recordset")
sql2="select * from person where uname='"&uname&"'"
rs.open sql2,conn,1,1
click=rs("click") %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" href="../inc/index.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>天天人才—&gt;个人登录首页</title>
</head>
<SCRIPT language=JavaScript src="../inc/window.js"></SCRIPT>
<form action=search.asp method=post>
<body topmargin="0" leftmargin="0">
<!--#include file="../inc/top2.htm"--> 
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="445">
    <tr>
      <td width="778" height="18" valign="top" colspan="5" bgcolor="#53BEB0"> 
        　 </td>
    </tr>
    <tr>
      <td width="139" valign="top" bgcolor="#53BEB0" height="371" rowspan="2"> 
        <p align="center">　 
        <div align="center">
          <center>
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
          </center>
        </div>
        <p align="center">&nbsp; </td>
      <td width="30" height="371" valign="top" rowspan="2"><img border="0" src="../images/selfk.GIF"></td>
  </center>
      <td width="471" height="349" valign="top">
      <div align="left">
        <table border="1" cellpadding="0" cellspacing="0" width="447" height="3" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
          <tr>
            <td width="445" height="8" valign="top">
 
            </td>
          </tr>
          <tr>
            <td width="445" height="109" valign="top" bgcolor="#EBEEF3" bordercolor="#000000" bordercolorlight="#C6CEDE" bordercolordark="#C6CEDE"><br>
              &nbsp;亲爱的用户:<%=uname%>，您好！<br>
              
			  <% if rs("job")<>"" then %> 
			  <br>
              &nbsp;您的求职简历共被查阅[<%=click%>]次！              
			  <%else%>
			  <br>
              &nbsp;您尚未登录个人简历，请尽快登录个人简历！                
			  <%end if%>
              <% if mailnum="0" then %>               
              <p>&nbsp;你的信箱内暂无新邮件！</p>

			  <%else%>
              <p>&nbsp;您的信箱内共有[<%=mailnum%>]封新邮件,请进入我的信箱查阅，回复信件！</p>
			  <%end if%>
          
          </td>
          </tr>
          <tr>
            <td width="445" height="3" valign="top">
 
            </td>
          </tr>
        </table>
      </div>
      <% set rs=server.createobject("adodb.recordset")
      sql3="select * from company where job<>'""' order by id desc"           
      rs.open sql3,conn,1,1%>
      <div align="left">
        <table border="1" cellpadding="0" cellspacing="0" width="450" height="24" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
          <tr>
            <td height="20" width="465" colspan="4" bgcolor="#C6CEDE" valign="bottom">&nbsp;最新十条招聘信息</td>
          </tr>
          <tr>
            <td height="19" width="187" bgcolor="#EBEEF3" valign="bottom">
              <p align="left">&nbsp;公司名称</td>
            <td height="19" width="204" bgcolor="#EBEEF3" valign="bottom">
              <p align="left">&nbsp;招聘职位</td>
            <td height="19" width="34" bgcolor="#EBEEF3" valign="bottom">
              <p align="center">发信</td>
            <td height="19" width="34" bgcolor="#EBEEF3" valign="bottom">
              <p align="center">收藏</td>
          </tr>
          <% do while not rs.eof %>
		  <tr>
   <td height="19" width="187" bgcolor="#EBEEF3" valign="bottom">&nbsp;<a href="javascript:openwin('../company.asp?uid=<%=rs("uname")%>','top=10,left=300,width=460,height=420')"><% if len(rs("cname"))>12 then%><%=Left(rs("cname"),12)%>...<%else%><%=rs("cname")%><%end if%></a></td>
   <td height="19" width="204" bgcolor="#EBEEF3" valign="bottom">&nbsp;<a 
   href="javascript:openwin('../job.asp?uid=<%=rs("uname")%>','top=10,left=300,width=460,height=420')"><%=rs("job")%>
   </a></td>
   <td height="19" width="34" bgcolor="#EBEEF3" valign="bottom">
    <p align="center"><a href="javascript:openwin('sendmail.asp?reid=<%=rs("uname")%>','top=20,left=200,width=470,height=400')"><font face="Wingdings" color="#000060" size="2">*</font></a></td>
   <td height="19" width="34" bgcolor="#EBEEF3" valign="bottom">
    <p align="center"><font face="Wingdings"><a href="addfav.asp?fav=<%=rs("uname")%>"><font color="#000060">1</font></a></font></td>
          </tr>
     <% c=c+1                                                                     
     rs.movenext                                                                     
     if c>=10 then exit do                                                                     
     loop                                                                    
     rs.close                                                                    
     set rs=nothing %>  
        </table>
      </div>
      </td>
  <center>
      <td width="1" height="371" valign="top" bgcolor="#00006A" rowspan="2"></td>
      <td width="126" height="371" valign="top" bgcolor="#F3F3F3" rowspan="2">　</td>
    </tr>
    <tr>
      <td width="471" height="22" valign="top">
      <p align="center"><font color="#000000"><br>
                职位搜索器: </font><INPUT       
                  maxLength=20 size=16 name=key style="background-color: #EBEBEB; color: #00006A; font-family: 宋体; font-size: 9pt" value="请输入关键字-S" onclick="Javascript:this.value='';">&nbsp;       
      <input type="submit" value="搜 索" style="font-family: 宋体; font-size: 9pt; color: #00006A">
      <br>
      <br>
      </td>
    </tr>
    <tr>
      <td width="778" height="1" valign="top" colspan="5" bgcolor="#53BEB0"> 
        <p align="center"> </td>
    </tr>
    <tr>
      <td height="3" valign="top" colspan="5" width="778">
      </td>
    </tr>
    <tr>
      <td height="14" valign="top" colspan="5" width="778">
      <p align="center"><script language="javascript" src="../inc/copyright.js"></script>
      </td>
    </tr>
    <tr>
      <td height="3" valign="top" colspan="5" width="778">
      </td>
    </tr>
  </table>
  </center>

</div>
</body>
</form>
</html>

