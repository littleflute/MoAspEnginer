<%if session("admin_pass")<>"ok" then
   response.redirect "adminlogin.asp"
   response.end
end if
user_id=request("user_id")
session("user_id")=user_id
response.redirect "your.asp"
%>
