
<!-- #include file="connect.asp" -->
<%rs.Open "Select top 8 * From data Order By time DESC",conn,3
'可将10修改为你所需要在首页显示的信息条数%>
<%
if not rs.eof then
	rs.movefirst
	dis_mes=dis_mes & ""
    do while not rs.eof
	dis_mes=dis_mes &"<img src=dot.gif> <a href=top/dismemo.asp?id=" & rs("id") & " target=_blank>"&rs("name") &"</a>["&rs("time") & "]</font><br>"
		rs.movenext
	loop
	dis_mes=dis_mes & ""
end if
rs.close
dis_mes="document.write('" & dis_mes & "');"
response.write dis_mes
%>