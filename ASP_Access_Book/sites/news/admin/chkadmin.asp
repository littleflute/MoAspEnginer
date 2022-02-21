<!--#include file=conn.asp-->
<%
	dim sql
	dim rs
	dim username
	dim password
	username=replace(trim(request("username")),"'","")
	password=replace(trim(Request("password")),"'","")
	set rs=server.createobject("adodb.recordset")
	sql="select * from admin where password='"&password&"' and username='"&username&"'"
	rs.open sql,conn,1,1
 	if not(rs.bof and rs.eof) then
 		if password=rs("password") then
			session("admin")=rs("username")
			session("flag")=rs("flag")
			Response.write"<script>"
			Response.write"parent.menu.location='left.asp';"
			Response.write"window.location='main.asp';"
			Response.write"</script>"
 		else
			call Error
 		end if
	else
		call Error()
	end if
	sub Error()
		response.write "   <html><head><link rel='stylesheet' href='style.css'></head><body>"
	    	response.write "   <br><br><br>"
	    	response.write "    <table align='center' width='300' border='0' cellpadding='4' cellspacing='0' class='border'>"
	    	response.write "      <tr > "
	    	response.write "        <td class='title' colspan='2' height='15'> "
	    	response.write "          <div align='center'>操作: 确认身份失败!</div>"
	    	response.write "        </td>"
	    	response.write "      </tr>"
	    	response.write "      <tr> "
	    	response.write "        <td class='tdbg' colspan='2' height='23'> "
	    	response.write "          <div align='center'><br><br>"
	    	response.write "      用户名或密码错误!!! <br><br>"
	    	response.write "        <a href='javascript:onclick=history.go(-1)'>返回</a>"        
	    	response.write "        <br><br></div></td>"
	    	response.write "      </tr>   </table></body></html>" 
	end sub
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
%>