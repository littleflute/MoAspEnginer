<%
'连接BBS数据库
dim connstr

set db=Server.CreateObject("ADODB.connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("bbsdata.mdb")
	db.Open connstr



%>

