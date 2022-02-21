<%
Function OpenOrGet_Database( SessionName )
   
Dim conn

   If Not IsObject(Session(SessionName)) Then
	  set conn=Server.CreateObject("ADODB.Connection")                                                                                          
                                                                                     
	  Conn.open "DBQ="+server.mappath("btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"                                       
      Set Session(SessionName) = conn
   End If
   Set OpenOrGet_Database = Session(SessionName)
End Function

Function OpenOrGet_RsAndPageSize( conn, sql, SessionName )
   Dim rs 

   If Not IsObject(Session(SessionName)) Then
      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open sql, conn, adOpenStatic,1,1
      Set Session(SessionName) = rs
      rs.PageSize = 6
   End If
   Set OpenOrGet_RsAndPageSize = Session(SessionName)
End Function

Function Open_RsAndPageSize( conn, sql, SessionName )
   Dim rs 

   Set rs = Server.CreateObject("ADODB.Recordset")
   rs.Open sql, conn, adOpenStatic,1,1
   Set Session(SessionName) = rs
   rs.PageSize = 6
   Set Open_RsAndPageSize = Session(SessionName)
End Function
%>
<html>

<head>
<meta http-equiv="Content-Type"
content="text/html; charset=gb_2312-80">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<title></title>
</head>

<body>
</body>
</html>
