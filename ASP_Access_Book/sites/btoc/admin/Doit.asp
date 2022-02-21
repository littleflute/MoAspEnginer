<%
	Flag = request("OP")
	Num =request("NumBer")
	set cn=Server.CreateObject("ADODB.Connection")

	cn.open "DBQ="+server.mappath("../btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"    

	if Flag = 1 then
		'作废
		SQL = "Update 订单管理 set 标记='废' where 订单号='" + Num + "'"
	else
		'完成
		SQL = "Update 订单管理 set 完成时间='" + FormatDateTime(now) + "' where 订单号='" + Num + "'"
	end if 
	cn.Execute SQL
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>入库操作</title>
<script language="javascript">
{
	alert("操作以执行完成!")
	window.close()
}
</script>
</head>

<body>

</body>

</html>
