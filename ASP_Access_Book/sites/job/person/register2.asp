<% Response.Buffer=True %>
<!--#include file="../inc/person.inc"-->
<!--#include file="../inc/html.inc"-->
<% uname=session("puid") 
   modify=request("modify")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from person where uname='"&uname&"'"
rs.open sql,conn,1,1 
if rs("iname")<>"" then
else 
response.write"<SCRIPT language=JavaScript>alert('�û��Ƿ��������밴��˳����д��ְ������');"
response.write"javascript:history.go(-1)</SCRIPT>"
end if %>
<% if modify<>"ture" and rs("job") <> "" then
   response.write"<SCRIPT language=JavaScript>alert('���Ѿ���¼���˼������벻Ҫ�ظ���¼��');"
   response.write"javascript:history.go(-1)</SCRIPT>"
   end if %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" href="../inc/register.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<% if modify<>"ture" then %>
<title>�����˲š�&gt;��¼��ְ����</title>
<%else%>
<title>�����˲š�&gt;������ְ����</title>
</head>
<%end if%>
<SCRIPT LANGUAGE="JavaScript">
<!--//
function check()
{
   if (document.register.gznum.value.length<1) 
		alert("�����빤������������");
   else if (isNaN(register.gznum.value))
		alert("����������ֻ����д���֣�");
   else if (document.register.gzjl.value.length<1) 
		alert("��������ϸ����������");
   else
		document.register.submit();
}
//-->
</SCRIPT>
<% if modify<>"ture" then %>
<FORM name=register action=register2.asp method=post>
<%else%>
<FORM name="register" action="register2.asp?modify=ture" method="post">
<%end if%>
<body topmargin="0" leftmargin="0">
<!--#include file="../inc/top2.htm"-->
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="340">
    <tr>
        <td width="778" height="18" valign="top" colspan="5" bgcolor="#53BEB0"> 
          �� </td>
    </tr>
    <tr>
        <td width="139" valign="top" bgcolor="#53BEB0"> 
          <p align="center">�� 
          <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="118" height="280">
            <tr>
              <td width="100%" height="163" background="../images/stat-bg.GIF" valign="top">
        <p align="center"><br>
        <a href="main.asp">��¼��ҳ</a><br>
        <br>
        <a href="register.asp">��¼��ְ����</a><br>
        <br>
        <a href="modify.asp">������ְ����</a><br>
        <br>
        <a href="../changepwd.asp?stype=person" target="_blank">�޸ĵ�¼����</a><br>
        <br><a href="search.asp">ȫ��ְλ�б�</a>
              </td>
            </tr>
            <tr>
              <td width="100%" height="117" background="../images/stat-bg.GIF" valign="top">
        <p align="center"><br>
        <br><a href="favorite.asp">�ҵ��ղؼ�</a><br>
        <br>
        <a href="mailbox.asp">�ҵ�����</a><br>
        <br>
        <a href="../exit.asp">�˳���¼</a>
              </td>
            </tr>
          </table>
        </div>
          <p align="center">&nbsp; </td>
      <td width="36" height="604" valign="top"><img border="0" src="../images/selfk.GIF"></td>
      <td width="478" height="248" valign="top">
  </center>
    
      ��
    
        <table border="1" cellpadding="0" cellspacing="0" width="93%" height="258" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
          <tr>
         <td width="100%" height="18" valign="bottom" bgcolor="#C6CEDE" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
              <p align="center">=== ������Ҫ�س� ===</td>                                              
          </tr>
          <tr>
            <td width="100%" height="124" valign="top" bgcolor="#F3F3F3">
              <p align="center"> 
              <br>
              �����س�:<SELECT 
                  size=1 name=language> <OPTION value=�� <%if rs("language") ="��" then Response.Write "selected"%>>��
                  <OPTION value=Ӣ�� <%if rs("language") ="Ӣ��" then Response.Write "selected"%>>Ӣ��
                  <OPTION value=���� <%if rs("language") ="����" then Response.Write "selected"%>>����
                  <OPTION value=���� <%if rs("language") ="����" then Response.Write "selected"%>>����
                  <OPTION value=���� <%if rs("language") ="����" then Response.Write "selected"%>>����
                  <OPTION value=���� <%if rs("language") ="����" then Response.Write "selected"%>>����
                  <OPTION value=������ <%if rs("language") ="������" then Response.Write "selected"%>>������
                  <OPTION value=�������� <%if rs("language") ="��������" then Response.Write "selected"%>>��������
                  <OPTION value=�������� <%if rs("language") ="��������" then Response.Write "selected"%>>��������
                  <OPTION value=���� <%if rs("language") ="����" then Response.Write "selected"%>>����</OPTION></SELECT>&nbsp;                                  
                  <SELECT size=1 name=lanlevel> 
                  <OPTION value=�� <%if rs("lanlevel") ="��" then Response.Write "selected"%>>��
                  <OPTION value=�ļ�  <%if rs("lanlevel") ="�ļ�" then Response.Write "selected"%>>�ļ�
                  <OPTION value=�˼�  <%if rs("lanlevel") ="�˼�" then Response.Write "selected"%>>�˼�
                  <OPTION value=����  <%if rs("lanlevel") ="����" then Response.Write "selected"%>>����
                  <OPTION value=����  <%if rs("lanlevel") ="����" then Response.Write "selected"%>>����
                  <OPTION value=��ͨ  <%if rs("lanlevel") ="��ͨ" then Response.Write "selected"%>>��ͨ
                  <OPTION value=����  <%if rs("lanlevel") ="����" then Response.Write "selected"%>>����
                  <OPTION value=һ��  
				  <%if rs("lanlevel") ="һ��" then Response.Write "selected"%>>һ��</OPTION></SELECT></p> 
              <p align="center">��ͨ���̶�:<SELECT 
                  size=1 name=pthua> 
                  <OPTION value=��׼ <%if rs("pthua") ="��׼" then Response.Write "selected"%>>��ͨ
                  <OPTION value=һ�� <%if rs("pthua") ="һ��" then Response.Write "selected"%>>һ��
                  <OPTION value=�ϲ� <%if rs("pthua") ="�ϲ�" then Response.Write "selected"%>>�ϲ�
                  </OPTION></SELECT>&nbsp; ���������:                                      
                  <SELECT size=1 name=computer> 
                  <OPTION value=һ�� <%if rs("computer") ="һ��" then Response.Write "selected"%>>һ��
                  <OPTION value=���� <%if rs("computer") ="����" then Response.Write "selected"%>>����
                  <OPTION value=���� <%if rs("computer") ="����" then Response.Write "selected"%>>����
                    <OPTION value=�ϲ� <%if rs("computer") ="�ϲ�" then Response.Write "selected"%>>�ϲ�</OPTION></SELECT></p> 
              <p align="center">������Ҫ�س�:<br>
              <font color="#FF0000">
              ����100�����ڣ�</font><br>
              <%if modify<>"ture" then 
			  kothertc=""
			  else
			  kothertc=replace(rs("othertc"),"<br>",chr(13))
              kothertc=replace(kothertc,"&nbsp;"," ")%>
			  <%end if%>          
              <textarea rows="6" name="othertc" cols="34"><%=kothertc%></textarea>          
              <br>
              <br>
			  </p> 
            </td>
          </tr>
          <tr>
            <td width="100%" height="18" valign="bottom" bgcolor="#C6CEDE" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
              <p align="center">=== ��ع������� ===                                     
            </td>
          </tr>
          <tr>
            <td width="100%" height="108" valign="top" bgcolor="#F3F3F3">
              <p align="center"><br>
              ��������:������ع������鹲��<INPUT 
                  maxLength=2 size=2 name=gznum value="<%=rs("gznum")%>">��</p>
              <p align="center">��ϸ��������:<BR>���ո�ʽ<font color="#FF0000">{ʼ(��.��) ��(��.��) ְ������ ��˾����}</font>��д         
			  <%if modify<>"ture" then 
			  kgzjl=""
			  else
			  kgzjl=replace(rs("gzjl"),"<br>",chr(13))
              kgzjl=replace(kgzjl,"&nbsp;"," ")%>
			  <%end if%>          
              <textarea rows="11" name="gzjl" cols="48"><%=kgzjl%></textarea>          
			  </p> 
              <p align="center">
			  <% if modify<>"ture" then %>
			  <input type="button" value="��һ��" onclick="javascript:history.go(-1)">  
              <input type="button" value="��һ��" onClick="check()"><br><br>  
              <%else%>
              <input type="button" value="�� ��" onClick="check()"> 
			  <br>
			  <br>
			  <%end if%>
            </td>
          </tr>
        </table>
      <% rs.close %>
      </td>
  <center>
      <td width="1" height="604" valign="top" bgcolor="#00006A"></td>
      <td width="119" height="604" valign="top" bgcolor="#F3F3F3">��</td>
    </tr>
    <tr>
        <td width="778" height="1" valign="top" colspan="5" bgcolor="#53BEB0"> 
          <p align="center"> </td>
    </tr>
    <tr>
      <td width="778" height="3" valign="top" colspan="5">
      <p align="center">
      </td>
    </tr>
    <tr>
      <td width="778" height="7" valign="top" colspan="5">
      <p align="center"><script language="javascript" src="../inc/copyright.js"></script>
      </td>
    </tr>
    <tr>
      <td width="778" height="3" valign="top" colspan="5">
      </td>
    </tr>
  </table>
  </center>
</div>
</form>
</body>

</html>
<% gzjl=htmlencode2(request("gzjl"))
if gzjl="" then Response.End
language=request("language")
lanlevel=request("lanlevel")
pthua=request("pthua")
computer=request("computer")
othertc=htmlencode2(request("othertc"))
gznum=request("gznum")
if othertc="" then othertc="�������س�" end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from person where uname='"&uname&"'"
rs.open sql,conn,3,3
rs("gznum")=gznum
rs("gzjl")=gzjl
rs("language")=language
rs("lanlevel")=lanlevel
rs("pthua")=pthua
rs("computer")=computer
rs("othertc")=othertc
rs.update
rs.close
if modify<>"ture" then
Response.Redirect "register3.asp" 
else
Response.Redirect "modify.asp" 
end if %>







