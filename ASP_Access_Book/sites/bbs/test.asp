<!--#include file="odbc_connection.asp"-->
<% =session("Adminflag") %><br>
<% If   session("Adminflag")= Request.QueryString("boardid") Then %>boardid <% Else %>����
<% End If %>
