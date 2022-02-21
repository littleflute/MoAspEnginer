<%
dim conn   
dim connstr
connstr="DBQ="+server.mappath("guest.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
conn.open connstr
'修改:
'--------------------------------------------------------------------
home="../"  '你的主页,引号不能去掉
Mypage =10  '设置每页显示多条留言
PassWord = "111"  '管理密码
%>

