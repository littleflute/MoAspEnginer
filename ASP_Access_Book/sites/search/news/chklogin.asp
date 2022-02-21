<%
dim sql
dim rs
dim seekerrs
dim founduser
dim username
dim companyid
dim password
dim errmsg
dim founderr
founderr=false
FoundUser=false
username=request.form("username")
password=request.Form("password")
if username="" then
   response.redirect "index.asp"
end if
if password="" then
     response.redirect "index.asp"
end if

 if username="admin" and password="admin" then
   response.cookies("adminok")=true
   response.redirect "manage.asp"
 else
     response.redirect "index.asp"
 end if
%>
