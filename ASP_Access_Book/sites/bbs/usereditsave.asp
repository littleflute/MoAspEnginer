<!-- #include file="conn.asp" -->
<%
id=request.form("id")
psw=request.form("psw")
pswc=request.form("pswc")
email=request.form("email")
name=request.form("name")
'**********����Ƿ���д�������������ǲ��Զ���������ҳ��
if id="" or psw="" or email="" or name="" then
errmsg=errmsg & "����дһ��Ҫ��**�����ݣ�\n����\n"
end if
if pswc<>psw  then
errmsg=errmsg & "������������벻ͬ��\n����\n"
end if
if errmsg<>"" then
    Conn.Close
    Set conn = nothing
    Set rsc = nothing
    response.write("<script>alert('" & errmsg & "');history.go(-1)</script>")
    response.end
end if
'**********������**********

homepage=request.form("homepage")
college=request.form("college")
qq=request.form("qq")
tel=request.form("tel")
addr=request.form("addr")
face=request.form("face")
head=request.form("head")

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from uers where id= '" & id & "'" 
rs.open sql, conn,1,3
rs("psw")=psw
rs("sex")=request.form("sex")
rs("qq")=qq
rs("email")=email
rs("homepage")=homepage
rs("huiyuan")=request.form("huiyuan")
rs("part")=request.form("part")
rs("name")=name
rs("class")=request.form("class")
rs("tel")=tel
rs("addr")=addr
rs("college")=college
rs("face")=face
rs("head")=head
rs.update
rs.close
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ɹ�������</title><meta http-equiv="refresh" content="3;URL=uerparticle.asp?id=<%=request.form("id")%>">
<style type="text/css">
<!--
a:visited{text-transform: none; text-decoration: none; font-style: normal; font-weight: normal; font-variant: normal; color: #000000}
a:hover{text-decoration:underline; color: #FF3333}
a:link{text-transform: none; text-decoration: none; font-weight: normal; font-variant: normal; color: #000000}
.p9 { font-size: 9pt; line-height: 13pt; text-decoration: none}
-->
</style>
</head>

<body>
<br>
<br>
<table width="348" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr class="p9"> 
    <td width="338" height="21" align="center" background="pic/h3.gif"><font color="#FFFFFF" size="+1"><strong>�����޸ĳɹ���</strong></font></td>
  </tr>
  <tr align="center" class="p9"> 
    <td>&nbsp;</td>
  </tr>
  <tr class="p9"> 
    <td height="17" align="center"><a href="uerparticle.asp?id=<%=request.form("id")%>">��ϲ�㣬��ɹ���½��������Զ�������ҳ</a></td>
  </tr>
  <tr> 
    <td height="18" align="center" class="p9"><a href="uerparticle.asp?id=<%=request.form("id")%>">����㲻��ȴ�����������</a></td>
  </tr>
  <tr> 
    <td height="18" align="center">&nbsp;</td>
  </tr>
</table>
</body>
</html>
