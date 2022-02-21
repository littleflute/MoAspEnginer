<%   set Conn=Server.CreateObject("ADODB.Connection")                                                  
                                             
	Conn.open "DBQ="+server.mappath("btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"   
%>