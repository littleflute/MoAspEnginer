
<%
	FF = request("Flag")
	Num = request("Num")
	if Num = "" then
		SQL = "Select * from �������� order by ������"
		set cn=Server.CreateObject("ADODB.Connection")

		cn.open "DBQ="+server.mappath("../btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 

		set rs = cn.Execute(SQL)
	else
		SQL = "Update �������� set ���='" + FF + "' where ������='" + Num + "'"
		set cn=Server.CreateObject("ADODB.Connection")

		cn.open "Provider=sqloledb;" & "Data Source=127.0.0.1;Initial Catalog=qiye_2;User Id=sa;Password=password;"  
		set rs = cn.Execute(SQL)
		SQL = "Select * from �������� order by ������"
		set rs = cn.Execute(SQL)
	end if
		

%>


<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>���������Ϲ��ﶩ������ϵͳ</title>
</head>

<body topmargin="0" leftmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="74%">
    <tr>
      <td width="100%" colspan="5">
        <hr>
      </td>
    </tr>
    <tr>
      <td width="20%" align="center"><font face="����Ҧ��" size="3" color="#0000FF"><a href="newform.asp">δ������</a></font></td>
      <td width="20%" align="center"><font face="����Ҧ��" size="3" color="#0000FF"><a href="oldform.asp">�Ѵ�����</a></font></td>
      <td width="20%" align="center"><font face="����Ҧ��" size="3" color="#0000FF"><a href="wasteform.asp">��Ч����</a></font></td>
      <td width="20%" align="center"><font face="����Ҧ��" size="3" color="#0000FF"><a href="formsearch.asp">������ѯ</a></font></td>
      <td width="20%" align="center"><font face="����Ҧ��" size="3" color="#0000FF"><a href="formmanage.asp">��������</a></font></td>
    </tr>
    <tr>
      <td width="100%" colspan="5">
        <hr><form name="Modi" action="formmanage.asp" method="POST"><input name="Flag" type="hidden"><input name="Num" type="hidden">
        <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#808080" height="39">
          <tr>
            <td width="3%" align="center" bgcolor="#FFFF66" height="19">&nbsp;</td>
            <td width="12%" align="center" bgcolor="#FFFF66" height="19"><font face="����" size="2"><b>��������</b></font></td>
            <td width="18%" align="center" bgcolor="#FFFF66" height="19"><font face="����" size="2"><b>�û�����</b></font></td>
            <td width="11%" align="center" bgcolor="#FFFF66" height="19"><font face="����" size="2"><b>���ͷ�ʽ</b></font></td>
            <td width="11%" align="center" bgcolor="#FFFF66" height="19"><font face="����" size="2"><b>���ʽ</b></font></td>
            <td width="19%" align="center" bgcolor="#FFFF66" height="19"><font face="����" size="2"><b>���ʱ��</b></font></td>
            <td width="18%" align="center" bgcolor="#FFFF66" height="19"><font face="����" size="2"><b>���ʱ��</b></font></td>
            <td width="8%" align="center" bgcolor="#FFFF66" height="19"><font face="����" size="2"><b>���</b></font></td>
          </tr>
<%
	Do while Not rs.EOF
	if rs.fields("���") = "��" then
    	FF = "��"
    else
    	FF = "��"
    end if  			
%>
          <tr>
            <td width="3%" align="center" bgcolor="#FFFFFF" height="16"><font size="2" color="blue"><a style="cursor: hand" href onClick="
            	Modi.Flag.value='<%=FF%>';
            	Modi.Num.value='<%=rs.fields("������")%>';
            	Modi.submit()
            	">��</a></font></td>
            <td width="12%" align="center" bgcolor="#FFCCFF" height="16"><font face="����" size="2"><%=rs.fields("������")%></font></td>
            <td width="18%" align="center" bgcolor="#CCFFFF" height="16"><font face="����" size="2"><%=rs.fields("�û���")%></font></td>
            <td width="11%" align="center" bgcolor="#FFCCFF" height="16"><font face="����" size="2"><%=rs.fields("����")%></font></td>
            <td width="11%" align="center" bgcolor="#CCFFFF" height="16"><font face="����" size="2"><%=rs.fields("����")%></font></td>
            <td width="19%" align="center" bgcolor="#FFCCFF" height="16"><font face="����" size="2"><%=rs.fields("��дʱ��")%></font></td>
            <td width="18%" align="center" bgcolor="#CCFFFF" height="16"><font face="����" size="2"><%=rs.fields("���ʱ��")%></font></td>
            <td width="8%" align="center" bgcolor="#FFCCFF" height="16"><font face="����" size="2"><%=rs.fields("���")%></font></td>
          </tr>
<%
	rs.Movenext
	loop
	cn.Close
%>
        </table></form>
        <p align="center">Copyright:������˾</td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>

