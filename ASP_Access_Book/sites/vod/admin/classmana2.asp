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
       Set cmdTemp = Server.CreateObject("ADODB.Command")
       set rs=server.createobject("adodb.recordset")
       cmdTemp.CommandText = "SELECT * FROM Nclass where Nclassid=" & request.form("subject")
       cmdTemp.CommandType = 1
       Set cmdTemp.ActiveConnection = conn   
       rs.Open cmdTemp, , 1, 3
       if err.Number<>0 then
            err.clear
            response.write " �� �� �� �� �� ʧ �� ! "
       else
         rs("Nclass") = trim(request.form("reTitle"))
         rs("classid") = request.form("classid")
         rs.Update
         rs.Close
         set rs=nothing
         set cmdTemp=nothing
         finished
       end if
     case "del"
	Rem ɾ������Ŀ
       sql="delete from Nclass where Nclassid=" & request.form("subject")
       conn.execute sql
	Rem ɾ���������Ŀ����
       'sql="delete from download where Nclassid=" & request.form("subject")
       'conn.execute sql
       if err.Number<>0 then
            err.clear
            response.write " �� �� �� �� �� ʧ �� ! "
       else
         finished
       end if
     case "new"
       Set cmdTemp = Server.CreateObject("ADODB.Command")
       set rs=server.createobject("adodb.recordset")
       cmdTemp.CommandText = "SELECT * FROM Nclass where (Nclassid IS NULL)"
       cmdTemp.CommandType = 1
       Set cmdTemp.ActiveConnection = conn   
       rs.Open cmdTemp, , 1, 3
       if err.Number<>0 or request.form("psubject")="" then
            err.clear
            response.write " �� �� �� �� �� ʧ �� ! "
       else
         rs.AddNew
           rs("Nclass") = request.form("newTitle")
           rs("classid") = request.form("psubject")
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
<p align="center"><b>ר���޸ĳɹ���</b></p>

<%
 end sub
%>