<%@ CODEPAGE = "936" %>
<!--#include file=conn.asp-->
<%
  	if session("admin")="" then
  		response.redirect "admin.asp"
  	end if
%>
<html>
<%
   Set rs=Server.CreateObject("Adodb.RecordSet")
   
   sql="select * from admin  order by id"
   
   rs.Open sql,conn,1,1
%>
<head>
<link rel="stylesheet" href="style.css">
</head>
<BODY bgcolor="#39867B">
<div align="center"></div>
<table border="0" cellspacing="1" width="758" align=center bgcolor="#C6EBDE">
  <tr>
    <td width="100%"><img src="../images/top.jpg" width="758" height="114"></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><br>
        <br>
        <table border="0" cellspacing="1" width="758" align=center bgcolor="#C6EBDE">
          <tr>
            <td width="100%">�˺ţ�<%=rs("id")%></td>
          </tr>
          <tr>
            <td width="100%" bgcolor="#FFFFFF">�ȼ���
                <%if rs("flag")=1 then%>
            �����û�
            <%end if%>
            <%if rs("flag")=2 then%>
            ��ͨ�û�
            <%end if%>
            <%if rs("flag")=3 then%>
            Ա��
            <%end if%>
            </td>
            <b>����Ա�ɽ��в���˵��</b>��<br>
            <br>
          1�����Ѿ��������޸Ļ�ɾ����������������ӽ��в����������û��������û�����ͨ����Ա <br>
          <br>
          2���������Ŀ��������޸�ɾ����������������ӽ��в����������û��������û� <br>
          <br>
          3���û��������޸�ɾ����������������ӽ��в����������û��������û� <br>
          <br>
          4��<font color=red>Ϊ��ϵͳ�İ�ȫ�ԣ��뿪���������˳�ϵͳ</font> <br>
          <br>
          </tr>
      </table></td>
  </tr>
</table>
</html>