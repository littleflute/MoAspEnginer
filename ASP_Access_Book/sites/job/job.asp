<% Response.Buffer=True %>
<!--#include file="inc/dbconn.inc"-->
<% uid=request("uid")
set rs=server.createobject("adodb.recordset")
sql1="update company set click=click+1 where uname='"&uid&"'"
rs.open sql1,conn,1,1
sql2="select * from company where job<>'""' and uname='"&uid&"'"
rs.open sql2,conn,1,1 %>
<% if rs.eof or rs.bof then
   response.write"<SCRIPT language=JavaScript>alert('对不起，该用户不存在或已被删除！');"
   response.write"javascript:window.close();</SCRIPT>" 
   end if %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/index.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>天天人才—&gt;人才市场—&gt;<%=rs("cname")%></title>
</head>
<div align="center">
  <table border="2" cellpadding="0" cellspacing="0" width="390" height="18" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
    <tr>
      <td width="382" height="19" bgcolor="#C6CEDE" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" valign="bottom" colspan="2">
        <p align="center">&nbsp;=== 招聘信息 ===</td>            
    </tr>
  <center>
    <tr>
  <td width="288" height="19" valign="bottom">&nbsp;公司名称:<a href="company.asp?uid=<%=rs("uname")%>"><%=rs("cname")%></a></td>
  </center>
      <td width="94" height="19" valign="bottom">
        <p align="right">点击数:<%=rs("click")%>&nbsp;</td>
    </tr>
  <center>
    <tr>
      <td width="384" height="19" bgcolor="#F3F3F3" valign="bottom" colspan="2">&nbsp;招聘职位:<%=rs("job")%></td>
    </tr>
    <tr>
      <td width="384" height="19" valign="bottom" colspan="2">&nbsp;招聘人数:[<%=rs("zpnum")%>]人</td>
    </tr>
    <tr>
      <td width="384" height="19" bgcolor="#F3F3F3" valign="bottom" colspan="2">&nbsp;工作地点:<%=rs("gzdd")%></td>
    </tr>
    <tr>
      <td width="384" height="19" valign="bottom" colspan="2">&nbsp;岗位描述:</td>
    </tr>
    <tr>
      <td height="3" valign="top" bgcolor="#00006A" width="384" colspan="2">
      </td>
    </tr>
    <tr>
      <td width="384" height="19" bgcolor="#F8F8F8" valign="bottom" colspan="2">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="379" align="right">
            <tr>
              <td width="377">&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("zptext")%></td>
            </tr>
          </table>
        </div>
        </td>
    </tr>
    <tr>
      <td height="3" valign="top" bgcolor="#00006A" width="384" colspan="2">
      </td>
    </tr>
    <tr>
      <td width="384" height="19" valign="bottom" colspan="2">&nbsp;相关要求:</td>
    </tr>
    <tr>
      <td height="3" valign="top" bgcolor="#00006A" width="384" colspan="2">
      </td>
    </tr>
    <tr>
      <td width="384" height="19" bgcolor="#F8F8F8" valign="bottom" colspan="2">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="379" align="right">
            <tr>
              <td width="377" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("xgyq")%></td>
            </tr>
          </table>
        </div>
        </td>
    </tr>
    <tr>
      <td height="3" valign="top" bgcolor="#00006A" width="384" colspan="2">
      </td>
    </tr>
  </center>
    <% if session("puid")<>"" then%>
	<tr>
      <td width="384" height="24" valign="bottom" colspan="2">
        <p align="center">[<a href="person/addfav.asp?fav=<%=rs("uname")%>">添加到收藏夹</a>]     
		[<a href="person/sendmail.asp?reid=<%=rs("uname")%>">发送站内信件</a>]</td> 
    </tr>
	<%end if%>
  </table>
</div>
<p align="center">
            【<a href="javascript:window.close()">关闭窗口</a>】</p>
<p align="center">　</p>
