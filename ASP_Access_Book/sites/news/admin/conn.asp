<%
dim conn
dim connstr
dim db
pathload=""
		dim startime
		dim pass_word,User_ID,Data_Source
		startime=timer()
		
		Set conn = Server.CreateObject("ADODB.Connection")
		connstr="DBQ="+server.mappath("/news.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
		conn.Open connstr

%>

