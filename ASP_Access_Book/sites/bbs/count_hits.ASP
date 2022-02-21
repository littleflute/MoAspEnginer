<%response.buffer=true%>
<!--#include file="odbc_connection.asp"-->
<%
dim bbs_id,boardid
bbs_id=request("bbs_id") 
boardid=Trim(Request.QueryString("boardid"))
'下面一段将点击次数加1，然后引导至particular.asp以显示内容
sql="update news set hits=hits+1 where bbs_id=" & bbs_id
db.execute(sql)
db.close
response.redirect "particular.asp?bbs_id=" & bbs_id &"&boardid="&boardid
%>
