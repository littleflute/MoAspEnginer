<% Response.Buffer=True %>
<!--#include file="../inc/admin.inc"-->
<!--#include file="../inc/html.inc"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" href="../inc/register.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�����˲š�&gt;�˲��г���&gt;�����¼</title>
</head>
<form action=addnews.asp method=post>
<body topmargin="0" leftmargin="0">

<BR>
  <div align="center">
    <center>
      <table border="1" cellpadding="0" cellspacing="0" width="450" height="320" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
        <tr>
          <td width="446" height="17" valign="bottom" background="../images/t-bg1.gif">
            <p align="center">=== ������� ===</td>                 
        </tr>
        <tr>
          
        <td valign="top" bgcolor="#fffff4" width="446"> 
          <p align="center"><br>
            �� ��:<input type="text" name="title" size="37" maxLength=25> 
            
          <p align="center">�� ��:<br>                
            <textarea rows="17" name="text" cols="52" ></textarea>
            <p align="center"><input type="submit" value="�� ��" style="position: relative; height: 20"><br>
            <br>
          </td>
        </tr>
      </table>
    </center>
  </div>
</body>

</html>
<% title=request("title")
text=htmlencode2(request("text"))
if title ="" or text="" then 
response.end
end if
set rs=server.createobject("adodb.recordset")
sql="select * from jobnews"
rs.open sql,conn,3,3
rs.addnew
rs("title")=title
rs("text")=text
rs("idate")=now()
rs.update
rs.close
set rs=nothing
response.write"<SCRIPT language=JavaScript>alert('������ӳɹ���');"
response.write"javascript:window.close();</SCRIPT>"%>
