<!--#include file="conn.asp"-->
<%
dim rs
dim formsize,formdata,bncrlf,divider,datastart,dataend,mydata

formsize=request.totalbytes
formdata=request.binaryread(formsize)
bncrlf=chrB(13) & chrB(10)
divider=leftB(formdata,clng(instrb(formdata,bncrlf))-1)
datastart=instrb(formdata,bncrlf & bncrlf)+4
dataend=instrb(datastart+1,formdata,divider)-datastart
mydata=midb(formdata,datastart,dataend)

Set rs = Server.CreateObject("ADODB.Recordset")
rs.open "select * from larchives where user_id=" & session("user_id"),conn,3,2
rs("photo")=rs("photo")+1
rs.update

%><!--#include file="connpic.asp"--><%

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "pic",conn,3,2
rs.addnew
rs("big").appendchunk mydata
rs("user_id")=session("user_id")
rs("date")=date
rs.update

set rs=nothing
set conn=nothing
response.redirect "sendphoto.asp"
%> 