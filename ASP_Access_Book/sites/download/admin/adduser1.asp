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
<HTML>
<HEAD>
<TITLE></TITLE>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel="stylesheet" type="text/css">
<SCRIPT language=javascript>

function check()

{
  if(document.form1.username.value=="")
    {
      alert("�û���Ϊ��");
      return false;
    }
    
  if(document.form1.newpin.value=="")
    {
      alert("���벻��Ϊ��");
      return false;
    }
    
  if((document.form1.newpin.value)!=(document.form1.re_newpin.value))
    {
      alert("���벻ƥ��");
      return false;
    }

}

</SCRIPT>
</HEAD>

<BODY bgcolor="#39867B" leftmargin="0" topmargin="0">
<DIV align="center">
  <P>&nbsp;</P>
  <FORM method="post" action="saveuser3.asp" name="form1" onsubmit="javascript:return check();">
    
		
    <TABLE width="40%" border="0" cellspacing="1" cellpadding="0" bgcolor="#E1F4EE" >
      <TR bgcolor="#CCCCCC"> 
        <TD height="25" bgcolor="#E1F4EE"> 
          <DIV align="center"><FONT size="2">������Ա</FONT></DIV>
        </TD>
      </TR>
      <TR> 
        <TD height="30"> 
          <DIV align="center"><FONT size="2">�� �� ���� 
            <INPUT type="text" name="user">
            </FONT></DIV>
        </TD>
      </TR>
      <TR> 
        <TD height="30"> 
          <DIV align="center"><FONT size="2">��ʼ���롡 
            <INPUT type="password" name="newpin">
            </FONT> </DIV>
        </TD>
      </TR>
      <TR> 
        <TD height="30"> 
          <DIV align="center"><FONT size="2">ȷ�����롡 
            <INPUT type="password" name="re_newpin">
            </FONT></DIV>
        </TD>
      </TR>
    </TABLE>
    <P>
      <INPUT type="submit" name="Submit" value="ȷ ��">
    </P>
  </FORM>
  <P>&nbsp;</P>
  <P align=center></P> </DIV>
</BODY>
</HTML>
