<% Response.Buffer=True %>
<!--#include file="inc/dbconn.inc"-->
<% uid=request("uid")
set rs=server.createobject("adodb.recordset")
sql="select * from company where job<>'""' and uname='"&uid&"'"
rs.open sql,conn,1,1 %>
<% if rs.eof or rs.bof then
   response.write"<SCRIPT language=JavaScript>alert('�Բ��𣬸��û������ڻ��ѱ�ɾ����');"
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
<title>�����˲š�&gt;<%=rs("cname")%></title>
</head>
<div align="center">
  <center>
  <table border="2" cellpadding="0" cellspacing="0" width="390" height="18" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
    <tr>
      <td width="388" height="19" bgcolor="#C6CEDE" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" valign="bottom">
        <p align="center">=== ��˾��ϸ���� ===</td>                 
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom">&nbsp;��˾����:<a href="job.asp?uid=<%=rs("uname")%>"><%=rs("cname")%></a></td>
    </tr>
    <tr>
      <td width="388" height="19" bgcolor="#F3F3F3" valign="bottom">&nbsp;������ҵ:<%=rs("trade")%></td>
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom">&nbsp;��ҵ����:<%=rs("cxz")%></td>
    </tr>
    <tr>
      <td width="388" height="19" bgcolor="#F3F3F3" valign="bottom">&nbsp;��������:<%=rs("area")%></td>
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom">&nbsp;��������:<%=rs("fdate")%></td>
    </tr>
    <tr>
      <td width="388" height="19" bgcolor="#F3F3F3" valign="bottom">&nbsp;ע���ʽ�:<%if rs("fund")="δ֪" then response.write"����"else response.write""&rs("fund")&"�������" end if%></td>
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom">&nbsp;ͨѶ��ַ:<%=rs("address")%>&nbsp;�ʱ�:<%=rs("zip")%></td>
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom" bgcolor="#F3F3F3">&nbsp;��ϵ��:<%=rs("pname")%></td>
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom">&nbsp;��ϵ�绰:<%=rs("phone")%></td>
    </tr>
    <tr>
      <td width="388" height="19" bgcolor="#F3F3F3" valign="bottom">&nbsp;�������:<%=rs("fax")%></td>
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom">&nbsp;��������:<a href="mailto:<%=rs("email")%>"><%=rs("email")%></a></td>
    </tr>
    <tr>
      <td width="388" height="19" bgcolor="#F3F3F3" valign="bottom">&nbsp;��˾��վ:<a href="<%=rs("http")%>"><%=rs("http")%></a></td>
    </tr>
    <tr>
      <td width="388" height="19" valign="bottom">&nbsp;��˾���:</td>
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
        <p align="center">[<a href="person/addfav.asp?fav=<%=rs("uname")%>">��ӵ��ղؼ�</a>]     
		[<a href="person/sendmail.asp?reid=<%=rs("uname")%>">����վ���ż�</a>]</td> 
    </tr>
	<%end if%>
  </table>
</div>
<p align="center">
            ��<a href="javascript:window.close()">�رմ���</a>��</p>
<p align="center">��</p>
