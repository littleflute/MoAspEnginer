<%
'定义数据库连接对象   
dim conn   
'定义数据库连接字符串
dim connstr

   on error resume next
   connstr ="DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & DefaultDir &server.mappath("/job.mdb")&";"

   set conn=server.createobject("ADODB.CONNECTION")
   if err then 
      err.clear
   else
        conn.open connstr 
        if err then 
           err.clear
        end if
   end if
%>
