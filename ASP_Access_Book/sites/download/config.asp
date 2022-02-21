<!-- #include file="conn1.asp" -->
<!-- #include file="inc/function.asp" -->
<%
dim Title_Name,CategoryName,CategoryName_CHS,username,URL,kicktime

'栏目变量名
CategoryName="SoftDown"   '不要改动

'网站名称
Title_Name="QQ软件" '可以改动

UserName= session("sUserName") '不要改动

'设置删除在线数据表中多少分钟内不活动用户,单位为分钟

kicktime=20

%>