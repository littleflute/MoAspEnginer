<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="conn.asp"-->
<%
  if session("admin")="" then
  response.redirect "admin.asp"
  else
	if session("flag")>1 then
		response.write "<br><p align=center>��û�в�����Ȩ��</p>"
		response.end
	end if
  end if

   dim sql
   dim rs
   dim cmdTemp
   
   select case request.form("options")
     case "rename"
       on error resume next
       Set cmdTemp = Server.CreateObject("ADODB.Command")
       set rs=server.createobject("adodb.recordset")
       cmdTemp.CommandText = "SELECT * FROM class where classid=" & request.form("subject")
       cmdTemp.CommandType = 1
       Set cmdTemp.ActiveConnection = conn   
       rs.Open cmdTemp, , 1, 3
       if err.Number<>0 then
            err.clear
            response.write " �� �� �� �� �� ʧ �� ! "
       else
         rs("class") = request.form("reTitle")
         rs.Update
         rs.Close
         set rs=nothing
         set cmdTemp=nothing
         finished
       end if
     case "del"
	Rem ɾ����Ŀ
        sql="delete from class where classid=" & request.form("subject")
        conn.execute sql
	Rem ɾ�������Ŀ������Ŀ
        sql="delete from Nclass where classid=" & request.form("subject")
        conn.execute sql
	Rem ɾ�������Ŀ��Ӧ����
        sql="delete from download where classid=" & request.form("subject")
        conn.execute sql
        if err.Number<>0 then
            err.clear
            response.write " �� �� �� �� �� ʧ �� ! "
        else
            finished
        end if
     case "new"
       on error resume next
       Set cmdTemp = Server.CreateObject("ADODB.Command")
       set rs=server.createobject("adodb.recordset")
       cmdTemp.CommandText = "SELECT * FROM class where (classid IS NULL)"
       cmdTemp.CommandType = 1
       Set cmdTemp.ActiveConnection = conn   
       rs.Open cmdTemp, , 1, 3
       if err.Number<>0 then
            err.clear
            response.write " �� �� �� �� �� ʧ �� ! "
       else
         rs.AddNew
         rs("class") = request.form("newTitle")
         rs.Update
         rs.Close
         set rs=nothing
         set cmdTemp=nothing
         finished
       end if
   end select

   sub finished()
     %>
<html>
<head>
<title>�޸ĳɹ�</title>
<meta http-equiv="refresh" content='0; URL=classmana.asp'>
</head>
<body>
<p align="center"><b>����޸ĳɹ���</b></p>

<%
 end sub
%>