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
ary_comp(0,1)="(FSO �ı��ļ���д)"
ary_comp(1,0)="Scripting.Dictionary"
ary_comp(1,1)="(Dictionary����)"
ary_comp(2,0)="ADODB.Connection"
ary_comp(2,1)="(ADO ���ݿ�֧��)"
ary_comp(3,0)="ADODB.Stream"
ary_comp(3,1)="(������ϴ�)"
ary_comp(4,0)="JRO.JetEngine"
ary_comp(4,1)="(���ݿ�ѹ��)"
ary_comp(5,0)="SoftArtisans.FileUp"
ary_comp(5,1)="(SA-FileUp �ļ��ϴ�)"
ary_comp(6,0)="LyfUpload.UploadFile"
ary_comp(6,1)="(LyfUpload �ļ��ϴ�)"
ary_comp(7,0)="JMail.Message"
ary_comp(7,1)="(Jmail �ʼ�֧��)"
ary_comp(8,0)="CDONTS.NewMail"
ary_comp(8,1)="(CDONTS �ʼ�֧��)"
%>
<table border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr bgcolor="#B4E7F1"> 
    <td width="45%" align="center" bgcolor="#B4E7F1">����</td>
					
    <td align="center" style="font-weight:600">����ֵ</td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">��������</td>
					<td><%=Request.ServerVariables("SERVER_NAME")%></td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">������IP</td>
					<td><%=Request.ServerVariables("HTTP_X_FORWARDED_FOR")%></td>
				</tr>
				<tr >
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">�������˿�</td>
					<td><%=Request.ServerVariables("SERVER_PORT")%></td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">������ʱ��</td>
					<td><%=now%></td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">IIS�汾</td>
					<td><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">�ű���ʱʱ��</td>
					<td><%=Server.ScriptTimeout%> ��</td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">session����ʱ��</td>
					<td><%=session.Timeout%> ��</td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">��ǰ�ļ�·��</td>
					<td><%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">������CPU����</td>
					<td><%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> ��</td>
				</tr>
				<tr class="rowbg2">
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">��������������</td>
					<td><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
				</tr>
				<tr>
					
    <td bgcolor="#B4E7F1" style="line-height:20px;">����������ϵͳ</td>
					<td><%=Request.ServerVariables("OS")%></td>
				</tr>
				<%  dim i,iszc 
				for i=0 to 8%>
				<tr>
    <td bgcolor="#B4E7F1">
<% if comp_check(ary_comp(i,0)) then iszc=true else iszc=false end if%> <%=ary_comp(i,1)%></td>
				<td><% if iszc=true then%>֧��<% else%>
      <font color="#FF0000">��֧��</font>
<% end if %></td> 
                   </tr>
			   <% next %>
			</table>