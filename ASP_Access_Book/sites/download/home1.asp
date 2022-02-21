<%sql="select top 1 * from home"
Set rs1= Server.CreateObject("ADODB.Recordset")
rs1.open sql,conn,0,1%>