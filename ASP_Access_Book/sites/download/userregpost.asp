<%PageName="UserRegPost"%>
<!--#include file="conn.asp"-->
<!--#include file="5324_lj.asp"-->
<!--#include file="INC/CHAR.INC"-->
<%
founderr=false

if request.form("User")="" or len(request.form("User"))>20 then
	errmsg=errmsg+"<br>"+"<li>�û����������(δ����򳤶ȳ�����20���ֽ�)��"
	founderr=true
else
	User=trim(request.form("User"))
end if
if request.form("sex")="" then
	errmsg=errmsg+"<br>"+"<li>��ѡ�������Ա�"
	founderr=true
elseif request.form("sex")=0 or request.form("sex")=1 then
	sex=request.form("sex")
else
	errmsg=errmsg+"<br>"+"<li>��������ַ��Ƿ���"
	founderr=true
end if
if request.form("password")="" or Len(request.form("password"))>20 then
	errmsg=errmsg+"<br>"+"<li>��������������(���Ȳ��ܴ���20)��"
	founderr=true
else
	password=request.form("password")
end if
if password<>request("password2") then
	errmsg=errmsg+"<br>"+"<li>������������ȷ�����벻һ�¡�"
	founderr=true
end if
if IsValidEmail(trim(request.form("Email")))=false then
	errmsg=errmsg+"<br>"+"<li>����Email�д���"
	founderr=true
else
	Email=trim(request.form("Email"))
end if



if founderr=true then
	call error()
else
	set rs=server.createobject("adodb.recordset")
	sql="select * from user where user='"&user&"'"
	rs.open sql,conn,1,3
	if not rs.eof or user=WebName then
		errmsg="<br>"+"<li>�Բ�����������û����Ѿ���ע�ᣬ���������롣"
		founderr=true
	else
		rs.addnew
		rs("user")=user
		rs("password")=password
		rs("email")=email
		if request.form("oicq")<>"" then rs("oicq")=request.form("oicq")
		Rs("sex")=sex
		rs("logins")=0
		rs("UserLevel")=1
		rs("UserPoint")=0
		rs.update
	end if
	rs.close

	if founderr=true then
		call error()
	else
%>
<!--#include file="head.asp"-->
<br>
<table cellpadding=0 cellspacing=0 border=0 width=730 bgcolor=#414141 align=center>
  <tr>
    <td align=center><br>
      <table border=1 cellpadding=3 cellspacing=1 border=0 class="TableLine" bordercolorlight="#414141" width=50% bgcolor=#414141>
        <TR align=middle bgcolor=#414141>
          <TD colSpan=2 height=24 bgcolor=#414141><b>��Աע��ɹ�</b></TD>
        </TR>
        <TR>
          <TD width=20% align=right>ע �� ��</TD>
          <TD><%=user%> ��</TD>
        </TR>
        <TR>
          <TD align=right>�� ��</TD>
          <TD><%if sex=1 then%>��<%else%>Ů<%end if%></TD>
        </TR>
        <TR>
          <TD align=right>�� ��</TD>
          <TD><%=password%>��</TD>
        </TR>
        <TR>
          <TD align=right>Email��ַ</TD>
          <TD><%=email%>��</TD></TR>
        <TR>
          <TD align=right>OICQ����</TD>
          <TD><%if request.form("oicq")="" then%>δע��<%else%><%=request.form("oicq")%><%end if%> </TD>
        </TR>
        <TR>
          <TD align=right>��Ա����</TD>
          <TD>0</TD>
        </TR>
        <TR>
          <TD align=right>��Ա�ȼ�</TD>
          <TD>��</TD>
        </TR>
              <TR>
          <TD align=center colSpan=2><a href=index.asp>���½</a></TD>
        </TR>
      </TABLE>
    <br></td>
  </tr>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
<%
	end if
	set rs=nothing
end if
conn.close
set conn=nothing
%>