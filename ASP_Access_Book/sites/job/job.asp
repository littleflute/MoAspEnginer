<% Response.Buffer=True %>
<!--#include file="inc/dbconn.inc"-->
<% uid=request("uid")
set rs=server.createobject("adodb.recordset")
sql1="update company set click=click+1 where uname='"&uid&"'"
rs.open sql1,conn,1,1
sql2="select * from company where job<>'""' and uname='"&uid&"'"
rs.open sql2,conn,1,1 %>
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
<title>�����˲š�&gt;�˲��г���&gt;<%=rs("cname")%></title>
</head>
<div align="center">
  <table border="2" cellpadding="0" cellspacing="0" width="390" height="18" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
    <tr>
      <td width="382" height="19" bgcolor="#C6CEDE" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" valign="bottom" colspan="2">
        <p align="center">&nbsp;=== ��Ƹ��Ϣ ===</td>            
    </tr>
  <center>
    <tr>
  <td width="288" height="19" valign="bottom">&nbsp;��˾����:<a href="company.asp?uid=<%=rs("uname")%>"><%=rs("cname")%></a></td>
  </center>
      <td width="94" height="19" valign="bottom">
        <p align="right">�����:<%=rs("click")%>&nbsp;</td>
    </tr>
  <center>
    <tr>
      <td width="384" height="19" bgcolor="#F3F3F3" valign="bottom" colspan="2">&nbsp;��Ƹְλ:<%=rs("job")%></td>
    </tr>
    <tr>
      <td width="384" height="19" valign="bottom" colspan="2">&nbsp;��Ƹ����:[<%=rs("zpnum")%>]��</td>
    </tr>
    <tr>
      <td width="384" height="19" bgcolor="#F3F3F3" valign="bottom" colspan="2">&nbsp;�����ص�:<%=rs("gzdd")%></td>
    </tr>
    <tr>
      <td width="384" height="19" valign="bottom" colspan="2">&nbsp;��λ����:</td>
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
      <td width="384" height="19" valign="bottom" colspan="2">&nbsp;���Ҫ��:</td>
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
        <p align="center">[<a href="person/addfav.asp?fav=<%=rs("uname")%>">��ӵ��ղؼ�</a>]     
		[<a href="person/sendmail.asp?reid=<%=rs("uname")%>">����վ���ż�</a>]</td> 
    </tr>
	<%end if%>
  </table>
</div>
<p align="center">
            ��<a href="javascript:window.close()">�رմ���</a>��</p>
<p align="center">��</p>
