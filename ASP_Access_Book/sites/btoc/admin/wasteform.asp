<%
	set cn=Server.CreateObject("ADODB.Connection")

	cn.open "DBQ="+server.mappath("../btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"   

	SQL = "Select * from �������� where ���='��' order by ������"
	set rs = cn.Execute(SQL)

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
        <hr>
        <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#808080">
          <tr>
            <td width="5%" align="center" bgcolor="#FFFF66">&nbsp;</td>
            <td width="19%" align="center" bgcolor="#FFFF66"><font face="����" size="2"><b>��������</b></font></td>
            <td width="22%" align="center" bgcolor="#FFFF66"><font face="����" size="2"><b>�û�����</b></font></td>
            <td width="17%" align="center" bgcolor="#FFFF66"><font face="����" size="2"><b>���ͷ�ʽ</b></font></td>
            <td width="17%" align="center" bgcolor="#FFFF66"><font face="����" size="2"><b>���ʽ</b></font></td>
            <td width="20%" align="center" bgcolor="#FFFF66"><font face="����" size="2"><b>��дʱ��</b></font></td>
          </tr>
<%
	Do while Not rs.EOF
%>
          <tr>
            <td width="5%" align="center" bgcolor="#FFFFFF"><font size="2" color="blue"><a href onClick="
            	window.open('formtxs.asp?Num=<%=rs.fields("������")%>&Nam=<%=rs.fields("�û���")%>','formtxt','width=375,height=400,scrollbars=1')" style="CURSOR:hand">.&gt;&gt;</a>
            	</font></td>
            <td width="19%" align="center" bgcolor="#FFCCFF"><font face="����" size="2"><%=rs.fields("������")%></font></td>
            <td width="22%" align="center" bgcolor="#CCFFFF"><font face="����" size="2"><%=rs.fields("�û���")%></font></td>
            <td width="17%" align="center" bgcolor="#FFCCFF"><font face="����" size="2"><%=rs.fields("����")%></font></td>
            <td width="17%" align="center" bgcolor="#CCFFFF"><font face="����" size="2"><%=rs.fields("����")%></font></td>
            <td width="20%" align="center" bgcolor="#FFCCFF"><font face="����" size="2"><%=rs.fields("��дʱ��")%></font></td>
          </tr>
<%
	rs.Movenext
	loop
	cn.Close
%>
        </table>
        <p align="center">Copyright:������˾</td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>
