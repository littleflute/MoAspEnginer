<% Response.Buffer=True %>
<!--#include file="inc/dbconn.inc"-->
<% uid=request("uid")
set rs=server.createobject("adodb.recordset")
sql1="update person set click=click+1 where uname='"&uid&"'"
rs.open sql1,conn,1,1
sql2="select * from person where job<>'""' and uname='"&uid&"'"
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
<title>�����˲š�&gt;<%=rs("iname")%></title>
</head>
<div align="center">
  <center>
  <table border="2" cellpadding="0" cellspacing="0" width="390" height="18" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
    <tr>
      <td height="19" bgcolor="#C6CEDE" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" valign="bottom" colspan="2" width="384">
        <p align="center">=== ������ְ���� ===</td>                                     
    </tr>
    <tr>
      <td height="19" valign="bottom" width="176">&nbsp;����:<%=rs("cname")%></td>
  </center>
   
      <td height="19" valign="bottom" width="206">
        <p align="right">�����:<%=rs("click")%>&nbsp;</p>
    </td>
    </tr>
  <center>
    <tr>
      <td height="19" valign="bottom" bgcolor="#EBEEF3" colspan="2" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="384">&nbsp;���˻�������    
        ------------------------------------------------&gt;</td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;�Ա�:<%=rs("sex")%>&nbsp;��������:<%=rs("bday")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">&nbsp;����:<%=rs("mzhu")%>&nbsp;������ò:<%=rs("zzmm")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;�������ڵ�:<%=rs("hka")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">&nbsp;��߽����̶�:<%=rs("edu")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;ר ҵ:<%=rs("zye")%></td>                     
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">&nbsp;��ҵԺУ:<%=rs("school")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;����ְ��:<%=rs("zchen")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#EBEEF3" colspan="2" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="384">&nbsp;ϣ��������������ϵ��Ϣ    
        --------------------------------------&gt;</td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;��ְ����:<%=rs("jobtype")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">&nbsp;ӦƸ��λ:<%=rs("job")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;ϣ�������ص�:<%=rs("gzdd")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">&nbsp;н��Ҫ��:<%if rs("yuex")="����" then response.write"����" else response.write"��н["&rs("yuex")&"]RMB" end if%></td>
    </tr>
    <% if rs("otheryq")<>"������Ҫ��" then%>
	<tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;����Ҫ��:</td>
    </tr>
    <tr>
      <td height="3" valign="top" bgcolor="#00006A" colspan="2" width="384">
      </td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="379" align="right">
            <tr>
              <td width="377">&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("otheryq")%></td>
            </tr>
          </table>
        </div>
        </td>
    </tr>
    <tr>
      <td height="3" valign="top" bgcolor="#00006A" colspan="2" width="384">
      </td>
    </tr>
    <%end if%>
	<tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;��ϵ��:<%=rs("cname")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">&nbsp;��ϵ�绰:<%=rs("phone")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;Ѱ��������:<%=rs("callnum")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">&nbsp;E-mail:<a href="mailto:<%=rs("email")%>"><%=rs("email")%></a></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;OICQ����:<%=rs("oicq")%></td>
    </tr>
    <tr>
  <td height="19" valign="bottom" colspan="2" bgcolor="#F3F3F3" width="384">&nbsp;������ҳ:<a href="<%=rs("http")%>"><%=rs("http")%></a></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;��ϵ��ַ:<%=rs("address")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#EBEEF3" colspan="2" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="384">&nbsp;������Ҫ�س�-��ع�������    
        -----------------------------------&gt;</td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;�����س�:<%=rs("language")%> �ȼ�:<%=rs("lanlevel")%></td>                 
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">&nbsp;��ͨ���̶�:<%=rs("pthua")%>&nbsp;���������:<%=rs("computer")%></td>   
    </tr>
    <% if rs("othertc")<>"�������س�" then%>
	<tr>
      <td height="19" valign="bottom" bgcolor="#FFFFFF" colspan="2" width="384">&nbsp;������Ҫ�س�:</td>
    </tr>
    <tr>
      <td height="3" valign="top" bgcolor="#00006A" colspan="2" width="384">
      </td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="379" align="right">
            <tr>
              <td width="377">&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("othertc")%></td>
            </tr>
          </table>
        </div>
        </td>
    </tr>
    <tr>
      <td height="3" valign="top" bgcolor="#00006A" colspan="2" width="384">
      </td>
    </tr>
  </center>
   <%end if%>
  <tr>
      <td height="19" valign="bottom" width="176">&nbsp;��ϸ��������:</td>
      <td height="19" valign="bottom" width="206">
        <p align="right">��ع�������:<%=rs("gznum")%>��&nbsp;&nbsp;</p>
      </td>
  </tr>
  <tr>
      <td height="3" valign="top" bgcolor="#00006A" colspan="2" width="384">
      </td>
  </tr>
  <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="379" align="right">
            <tr>
              <td width="377">&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("gzjl")%></td>
            </tr>
          </table>
        </div>
        </td>
  </tr>
  <tr>
      <td height="3" valign="top" bgcolor="#00006A" colspan="2" width="384">
      </td>
  </tr>
	 <% if session("cuid")<>"" then%>
	 <tr>
      <td height="24" valign="bottom" colspan="2" width="384">
        <p align="center">[<a href="company/addfav.asp?fav=<%=rs("uname")%>">��ӵ��ղؼ�</a>]                        
		[<a href="company/sendmail.asp?reid=<%=rs("uname")%>">����վ���ż�</a>]</td> 
    </tr>
	<%end if%>
  </table>
</div>
<br>
<CENTER>��<a href="javascript:window.close()">�رմ���</a>��</CENTER>
</html>

