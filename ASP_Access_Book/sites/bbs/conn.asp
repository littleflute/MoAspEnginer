<%
   dim conn   
   dim connstr
   on error resume next

set conn=Server.CreateObject("ADODB.connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("bbsdata.mdb")
	conn.Open connstr
%>