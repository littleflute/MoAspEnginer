<!-- #include file="conn.asp" -->
<%

from=request.form("from")
content=request.form("content")
title=request.form("title")
ttto= request("to")
'**********����Ƿ���д�������������ǲ��Զ���������ҳ��
if  ttto=""  or  title ="" or  content=""  then
errmsg=errmsg & "����дһ��Ҫ��**�����ݣ�\n����\n"
end if
'**********����û���.������ظ��û������Զ���������ҳ��
dim rsc,errmsg
Set rsc = Conn.Execute("select * from uers where id= '" & ttto & "'")
if    rsc.eof  then  errmsg=errmsg & "û���ҵ����û���\n"

if errmsg<>"" then
    Conn.Close
    Set conn = nothing
    Set rsc = nothing
    response.write("<script>alert('" & errmsg & "');history.go(-1)</script>")
    response.end
end if
'**********������**********
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from duanx"
rs.open sql, conn,1,3
rs.addnew
rs("to")=ttto
rs("from")=from
rs("title")=title
rs("content")=content
rs("flag")="��"
rs.update
rs.close
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ɹ�������</title><meta http-equiv="refresh" content="1;URL=duanadd.asp">

</head>

<body>
<div align="center">
  <p><font color="#FF0000" size="3"><strong>�ɹ�ע�ᣡ</strong></font></p>
  <p>ϵͳ����3�����Զ��������������ϣ�</p>
  <p></p>
  <p><a href="duanadd.asp">����㲻������ȴ���������<br>
    (�������������޷��Զ�����) <br>
    </a> </p>
</div>
</body>
</html>
