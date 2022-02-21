<% Response.Buffer=True %>
<!--#include file="inc/dbconn.inc"-->
<% uid=request("uid")
set rs=server.createobject("adodb.recordset")
sql="select * from company where job<>'""' and uname='"&uid&"'"
rs.open sql,conn,1,1 %>
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
<title>天天人才—&gt;<%=rs("cname")%></title>
</head>
<div align="center">
  <center>
  <table border="2" cellpadding="0" cellspacing="0" width="390" height="18" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
    <tr>
      <td width="388" height="19" bgcolor="#C6CEDE" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" valign="bottom">
        <p align="center">=== 公司详细资料 ===</td>                 
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom">&nbsp;公司名称:<a href="job.asp?uid=<%=rs("uname")%>"><%=rs("cname")%></a></td>
    </tr>
    <tr>
      <td width="388" height="19" bgcolor="#F3F3F3" valign="bottom">&nbsp;所属行业:<%=rs("trade")%></td>
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom">&nbsp;企业性质:<%=rs("cxz")%></td>
    </tr>
    <tr>
      <td width="388" height="19" bgcolor="#F3F3F3" valign="bottom">&nbsp;所在区域:<%=rs("area")%></td>
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom">&nbsp;成立日期:<%=rs("fdate")%></td>
    </tr>
    <tr>
      <td width="388" height="19" bgcolor="#F3F3F3" valign="bottom">&nbsp;注册资金:<%if rs("fund")="未知" then response.write"面议"else response.write""&rs("fund")&"万人民币" end if%></td>
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom">&nbsp;通讯地址:<%=rs("address")%>&nbsp;邮编:<%=rs("zip")%></td>
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom" bgcolor="#F3F3F3">&nbsp;联系人:<%=rs("pname")%></td>
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom">&nbsp;联系电话:<%=rs("phone")%></td>
    </tr>
    <tr>
      <td width="388" height="19" bgcolor="#F3F3F3" valign="bottom">&nbsp;传真号码:<%=rs("fax")%></td>
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom">&nbsp;电子信箱:<a href="mailto:<%=rs("email")%>"><%=rs("email")%></a></td>
    </tr>
    <tr>
      <td width="388" height="19" bgcolor="#F3F3F3" valign="bottom">&nbsp;公司网站:<a href="<%=rs("http")%>"><%=rs("http")%></a></td>
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom">&nbsp;公司简介:</td>
    </tr>
    <tr>
      <td height="3" valign="top" bgcolor="#00006A">
      </td>
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom" bgcolor="#F3F3F3">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="379" align="right">
            <tr>
              <td width="377">&nbsp;&nbsp;<%=rs("jianj")%></td>
            </tr>
          </table>
        </div>
        </td>
    </tr>
    <tr>
      <td height="3" valign="top" bgcolor="#00006A">
      </td>
    </tr>
  </center>
    <% if session("puid")<>"" then%>
	<tr>
      <td width="388" height="24" valign="bottom">
        <p align="center">[<a href="person/addfav.asp?fav=<%=rs("uname")%>">添加到收藏夹</a>]     
		[<a href="person/sendmail.asp?reid=<%=rs("uname")%>">发送站内信件</a>]</td> 
    </tr>
	<%end if%>
  </table>
</div>
<p align="center">
            【<a href="javascript:window.close()">关闭窗口</a>】</p>
<p align="center">　</p>
