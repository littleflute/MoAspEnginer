<!--#include file="odbc_connection.asp"-->
<% =session("Adminflag") %><br>
<% If   session("Adminflag")= Request.QueryString("boardid") Then %>boardid <% Else %>²»ÐÐ
<% End If %>
