<%
dim server_v1,server_v2
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
   Response.write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title>出现错误！</title><div align=""center"" style=""width:400px;height:100px;padding: 8px;font-size:9pt;border: 1px solid ThreeDShadow;POSITION:absolute;top:expression((document.body.offsetHeight-100)/2);left:expression((document.body.offsetWidth-400)/2);""><table width=""100%"" height=""100%"" style=""font-size:12px;""><tr><td align=""center"">你提交的路径有误，禁止从站点外部提交数据请不要乱该参数！</td></tr></table>"
   Response.end
end if
%>