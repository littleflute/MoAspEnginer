<!--#include file=conn.asp-->
<%
	dim sql
	dim rs
	dim username
	dim password
	username=replace(trim(request("username")),"'","��")
	password=replace(trim(Request("password")),"'","��")

	set rs=server.createobject("adodb.recordset")
	sql="select * from admin where password='"&password&"' and username='"&username&"'"
	rs.open sql,conn,1,1
 	if not(rs.bof and rs.eof) then
 		if password=rs("password") then
			session("admin")=rs("username")
			session("flag")=rs("flag")
			Response.Redirect "manage.asp"
 		else
			call Error
 		end if
	else
		call Error()
	end if

	sub Error()
response.write " <LINK rel='stylesheet' type='text/css' href='style.css'> "
response.write " <TITLE>����Ա�����֤</TITLE> "
response.write " <BODY bgcolor='#009458'> "
response.write " <BR><BR><BR>  "   
response.write " <TABLE align='center' width='300' cellpadding='1' cellspacing='1'> "      
response.write " <TR bgcolor='#CCCCCC'>  "
response.write " <TD colspan='2' height='15' bgcolor='#145f74'> "     
response.write " <DIV align='center'><FONT color='#FFFFFF'>������ȷ�����ʧ�ܣ�</FONT></DIV> "
response.write " </TD> "
response.write " </TR> "   
response.write " <TR> "
response.write " <TD colspan='2' height='23' bgcolor='#a5d0dc'> "   
response.write " <DIV align='center'><BR><BR>�Ƿ���½�����Ĳ����Ѿ�����¼������<BR> "
response.write " <BR><A href='login.asp'>�ٴε�¼��</A><BR> "
response.write " <BR> "
response.write " </DIV> "
response.write " </TD> "
response.write " </TR> " 
response.write " </TABLE> " 
	end sub
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing

%>

