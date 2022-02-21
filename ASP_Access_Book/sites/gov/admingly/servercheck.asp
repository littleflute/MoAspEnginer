<!--#include file="admintop.asp"-->
<%
Function comp_check(byval str_class)
	on error resume next
	dim obj_check
	set obj_check = server.createobject(str_class)
	set obj_check=nothing
	if err.number<>0 then
		comp_check=false
	else
		comp_check=true
	end if
	err.clear()
End Function
dim ary_comp(8,1)
ary_comp(0,0)="Scripting.FileSystemObject"
ary_comp(0,1)="(FSO 文本文件读写)"
ary_comp(1,0)="Scripting.Dictionary"
ary_comp(1,1)="(Dictionary对象)"
ary_comp(2,0)="ADODB.Connection"
ary_comp(2,1)="(ADO 数据库支持)"
ary_comp(3,0)="ADODB.Stream"
ary_comp(3,1)="(无组件上传)"
ary_comp(4,0)="JRO.JetEngine"
ary_comp(4,1)="(数据库压缩)"
ary_comp(5,0)="SoftArtisans.FileUp"
ary_comp(5,1)="(SA-FileUp 文件上传)"
ary_comp(6,0)="LyfUpload.UploadFile"
ary_comp(6,1)="(LyfUpload 文件上传)"
ary_comp(7,0)="JMail.Message"
ary_comp(7,1)="(Jmail 邮件支持)"
ary_comp(8,0)="CDONTS.NewMail"
ary_comp(8,1)="(CDONTS 邮件支持)"
%>
<table border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr bgcolor="#B4E7F1"> 
    <td width="45%" align="center" bgcolor="#B4E7F1">名称</td>
					
    <td align="center" style="font-weight:600">参数值</td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">服务器名</td>
					<td><%=Request.ServerVariables("SERVER_NAME")%></td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">服务器IP</td>
					<td><%=Request.ServerVariables("HTTP_X_FORWARDED_FOR")%></td>
				</tr>
				<tr >
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">服务器端口</td>
					<td><%=Request.ServerVariables("SERVER_PORT")%></td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">服务器时间</td>
					<td><%=now%></td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">IIS版本</td>
					<td><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">脚本超时时间</td>
					<td><%=Server.ScriptTimeout%> 秒</td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">session过期时间</td>
					<td><%=session.Timeout%> 分</td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">当前文件路径</td>
					<td><%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">服务器CPU数量</td>
					<td><%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 个</td>
				</tr>
				<tr class="rowbg2">
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">服务器解译引擎</td>
					<td><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">服务器操作系统</td>
					<td><%=Request.ServerVariables("OS")%></td>
				</tr>
				<%  dim i,iszc 
				for i=0 to 8%>
				<tr>
    <td bgcolor="#B4E7F1">
<% if comp_check(ary_comp(i,0)) then iszc=true else iszc=false end if%> <%=ary_comp(i,1)%></td>
				<td><% if iszc=true then%>支持<% else%>
      <font color="#FF0000">不支持</font>
<% end if %></td> 
                   </tr>
			   <% next %>
			</table>