<%@language=vbscript%>
<!--#include file="conn.asp"-->
<%
  if session("admin")="" then
  response.redirect "admin.asp"
  else
	if session("flag")>1 then
		response.write "<br><p align=center>��û�в�����Ȩ��</p>"
		response.end
	end if
  end if

dim rs, sql
%>
<html>
<head>
<title>�û�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<%
   Set rs=Server.CreateObject("Adodb.RecordSet")
   
   sql="select * from admin where flag>="&Session("flag")
   
   rs.Open sql,conn,1,1
%>
<body>
<div align="center">
  <p>��</p>
  <p>�޸Ĺ���Ա��Ϣ | <a href=adduser.asp>���ӹ���Ա</a></p>
  <table width="501" border="0" cellspacing="1" cellpadding="0">
    <tr class="title"> 
      <td height="25"  width="30"> 
        <div align="center">���</div>
      </td>
      <td  width="120"> 
        <div align="center">�û���</div>
      </td>
      <td  width="120"> 
        <div align="center">����</div>
      </td>
      <td  width="100"> 
        <div align="center">Ȩ��</div>
      </td>
      <td  width="130"> 
        <div align="center">����</div>
      </td>
    </tr>
  </table>
  <%while not rs.EOF %>
  <%if (rs("flag")>Session("flag")) or (rs("username")=Session("admin")) then%>
  <form method="post" action="saveuser.asp" style="margin:0">
    <table width="500" border="0" cellspacing="1" cellpadding="0">
      <tr class="tdbg"> 
        <td height="30" width="30"> 
          <div align="center">1</div>
        </td>
        <td height="30" width="120"> 
          <div align="center">
            <input class="smallinput" type="text" name="manager" value="<%=rs("username")%>" size="12">
          </div>
        </td>
        <td height="30" width="120"> 
          <div align="center">
            <input class="smallinput" type="password" name="newpin" value="<%=rs("password")%>" size="12">
          </div>
        </td>
        <td height="30" width="100">
          <%if rs("flag")=1 then%> 
          <div align="center">�����û�</div>
          <%end if%>
          <%if rs("flag")=2 then%> 
          <div align="center">��ͨ�û�</div>
          <%end if%>
          <%if rs("flag")=3 then%> 
          <div align="center">Ա��</div>
          <%end if%>
        </td>
        <td height="30" width="130"> 
          <div align="center">
            <input class="buttonface" type="submit" name="Submit" value="�޸�">&nbsp;
            <input class="buttonface" type="submit" name="Submit" value="ɾ��">
            <input type="hidden" name="oldmanager" value="<%=rs("username")%>">
            <input type="hidden" name="oldpin" value="<%=rs("password")%>">
          </div>
        </td>
      </tr>
    </table>
  </form>
  <%else%>
    <table width="500" border="0" cellspacing="2" cellpadding="0">
      <tr bgcolor="#EBEBEB"> 
        <td height="30" width="30"> 
          <div align="center">1</div>
        </td>
        <td height="30" width="120"> 
          <div align="center">
            <%=rs("username")%>
          </div>
        </td>
        <td height="30" width="120"> 
          <div align="center">
            ******
          </div>
        </td>
        <td height="30" width="100"> 
          <%if rs("flag")=1 then%> 
          <div align="center">�����û�</div>
          <%else%>
          <div align="center">��ͨ�û�</div>
          <%end if%>
        </td>
        <td height="30" width="130"> 
          <div align="center">&nbsp;
          </div>
        </td>
      </tr>
    </table>  
    
  <%
     end if
     
     rs.MoveNext
     
     Wend
  %>
  <p>��</p> </div>
<%
  rs.Close
  set rs=Nothing
  
  conn.Close
  set conn=Nothing
%>
<p align=center>�����û����Խ������й��ܵĲ���<br>��ͨ�û�û���û��������Ŀ����Ȩ�ޣ��û�ֻ����ӳ����Ȩ�ޡ�
</body>
</html>