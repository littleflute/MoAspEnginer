<%
	set cn=Server.CreateObject("ADODB.Connection")

	cn.open "DBQ="+server.mappath("../btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"  
	SQL = "Select * from �������� where ������='" + request("Num") + "'"
	set rs = cn.Execute(SQL)
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>��������</title>
</head>

<body topmargin="0" leftmargin="0">

<table border="0" cellpadding="0" cellspacing="0" width="353" height="236">
  <tr>
    <td width="20" rowspan="4" height="96"></td>
    <td width="233" height="24"></td>
    <td width="110" height="99" rowspan="4"><img border="0" src="PE02097_.wmf" width="125" height="99"></td>
  </tr>
  <tr>
    <td width="233" height="26">&nbsp;<font size="3" face="����Ҧ��" color="#000080">��������:&nbsp;<%=request("Num")%></font></td>
  </tr>
  <tr>
    <td width="233" height="21">&nbsp;<font size="3" face="����Ҧ��" color="#000080">�û��ʺ�:&nbsp;<%=request("Nam")%></font></td>
  </tr>
  <tr>
    <td width="233" height="28"></td>
  </tr>
  <tr>
    <td width="363" colspan="3" height="9">
      <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#808080">
        <tr>
          <td width="197" bgcolor="#C0C0C0" align="center"><font face="����Ҧ��" size="3">��Ʒ����</font></td>
          <td width="83" bgcolor="#C0C0C0" align="center"><font face="����Ҧ��" size="3">�۸�</font></td>
          <td width="71" bgcolor="#C0C0C0" align="center"><font face="����Ҧ��" size="3">����</font></td>
        </tr>
<%
	dim TotalMoney
	Do while Not rs.EOF
%>
        <tr>
          <td width="197"><font size="2"><%=rs.fields("��Ʒ")%></font></td>
          <td width="83">
            <p align="center"><font size="2"><%=rs.fields("�۸�")%></font></td>
          <td width="71">
            <p align="center"><font size="2"><%=rs.fields("����")%></font></td>
        </tr>
<%
	TotalMoney = TotalMoney + clng(rs.fields("�۸�")) * clng(rs.fields("����"))
	rs.Movenext
	Loop
	
	SQL = "Select * from �ͻ���Ϣ where �˺�='" + request("Nam") + "'"
	set rs = cn.Execute(SQL)
%>      <tr>
			<td colspan="3">
              <p align="right"><font face="����Ҧ��">�����ܼ�Ϊ:<font color="#FF0000"><%=FormatNumber(TotalMoney,2)%></font><font color="#000000">Ԫ</font></font>
            </td>
		</tr>
		</table>
    </td>
  </tr>
  <tr>
    <td width="363" colspan="3" height="19" valign="top">
      <p align="right">
    </td>
  </tr>
  <tr>
    <td width="363" colspan="3" height="55">
      <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#808080">
        <tr>
          <td width="73" align="center" bgcolor="#C0C0C0"><font size="2">����</font></td>
          <td width="109"><font size="2"><%=rs.fields("����")%></font></td>
          <td width="74" align="center" bgcolor="#C0C0C0"><font size="2">�绰</font></td>
          <td width="101"><font size="2"><%=rs.fields("�绰")%></font></td>
        </tr>
        <tr>
          <td width="72" align="center" bgcolor="#C0C0C0"><font size="2">��ַ</font></td>
          <td width="281" colspan="3"><font size="2"><%=rs.fields("סַ")%></font></td>
        </tr>
        <tr>
          <td width="72" align="center" bgcolor="#C0C0C0"><font size="2">�ʼ�</font></td>
          <td width="281" colspan="3"><font size="2"><%=rs.fields("�ö�")%></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<% cn.close %>
</body>

</html>
