<%@language=vbscript%>
<%
  if session("admin")="" then
  response.redirect "admin.asp"
  else
	if session("flag")>1 then
		response.write "<br><p align=center>��û�в�����Ȩ��</p>"
		response.end
	end if
  end if

%>
<!--#include file="conn.asp"-->
<HTML>
<HEAD>
<TITLE>�û�����</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<%
   Set rs=Server.CreateObject("Adodb.RecordSet")
   
   sql="select * from [user]  order by id"
   
   rs.Open sql,conn,1,1
%>
<BODY bgcolor="#009458" leftmargin="0" topmargin="0">
<DIV align="center">
  <P>&nbsp;</P>
  
  <P><FONT color="#FFFFFF">�޸Ļ�Ա��Ϣ</FONT> | <A href=adduser1.asp>���ӻ�Ա</A></P>
  
	
  <TABLE width="500" border="0" cellspacing="1" cellpadding="0">
    <TR bgcolor="#145f74"> 
      <TD height="30" width="30"> 
        <DIV align="center"><FONT color="#FFFFFF">���</FONT></DIV>
      </TD>
      <TD height="30" width="120"> 
        <DIV align="center"><FONT color="#FFFFFF">�û���</FONT></DIV>
      </TD>
      <TD height="30" width="120"> 
        <DIV align="center"><FONT color="#FFFFFF">����</FONT></DIV>
      </TD>
      <TD height="30" bgcolor="#145f74"> 
        <DIV align="center"></DIV>
        <DIV align="center"><FONT color="#FFFFFF">����</FONT></DIV>
      </TD>
    </TR>
  </TABLE>
  <%while not rs.EOF %>
  
  <FORM method="post" action="saveuser2.asp" style="margin:0">
    
		
    <TABLE width="500" border="0" cellspacing="1" cellpadding="0">
      <TR bgcolor="#eeeeee"> 
        <TD height="30" width="30" bgcolor="#a5d0dc"> 
          <DIV align="center">1</DIV>
        </TD>
        <TD height="30" width="120" bgcolor="#cce6ed"> 
          <DIV align="center"> 
            <INPUT type="text" name="manager" value="<%=rs("user")%>" size="12">
          </DIV>
        </TD>
        <TD height="30" width="120" bgcolor="#a5d0dc"> 
          <DIV align="center"> 
            <INPUT type="password" name="newpin" value="<%=rs("password")%>" size="12">
          </DIV>
        </TD>
        <TD height="30" bgcolor="#cce6ed"> 
          <DIV align="center"></DIV>
          <DIV align="center"> 
            <INPUT type="submit" name="Submit" value="�޸�">
            &nbsp; 
            <INPUT type="submit" name="Submit" value="ɾ��">
            <INPUT type="hidden" name="oldmanager" value="<%=rs("user")%>">
            <INPUT type="hidden" name="oldpin" value="<%=rs("password")%>">
          </DIV>
        </TD>
      </TR>
    </TABLE>
  </FORM>
  <%
     
     
     rs.MoveNext
     
     Wend
  %>
  <P>&nbsp;</P> </DIV>
<%
  rs.Close
  set rs=Nothing
  
  conn.Close
  set conn=Nothing
%>
<P align=center>&nbsp; 
</BODY>
</HTML>
