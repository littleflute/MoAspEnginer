<%
	Flag = request("OP")
	Num =request("NumBer")
	set cn=Server.CreateObject("ADODB.Connection")

	cn.open "DBQ="+server.mappath("../btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"    

	if Flag = 1 then
		'����
		SQL = "Update �������� set ���='��' where ������='" + Num + "'"
	else
		'���
		SQL = "Update �������� set ���ʱ��='" + FormatDateTime(now) + "' where ������='" + Num + "'"
	end if 
	cn.Execute SQL
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>������</title>
<script language="javascript">
{
	alert("������ִ�����!")
	window.close()
}
</script>
</head>

<body>

</body>

</html>
