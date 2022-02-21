<%PageName="UserRegPost"%>
<!--#include file="conn.asp"-->
<!--#include file="5324_lj.asp"-->
<!--#include file="INC/CHAR.INC"-->
<%
founderr=false

if request.form("User")="" or len(request.form("User"))>20 then
	errmsg=errmsg+"<br>"+"<li>用户名输入错误(未输入或长度超过了20个字节)。"
	founderr=true
else
	User=trim(request.form("User"))
end if
if request.form("sex")="" then
	errmsg=errmsg+"<br>"+"<li>请选择您的性别。"
	founderr=true
elseif request.form("sex")=0 or request.form("sex")=1 then
	sex=request.form("sex")
else
	errmsg=errmsg+"<br>"+"<li>您输入的字符非法。"
	founderr=true
end if
if request.form("password")="" or Len(request.form("password"))>20 then
	errmsg=errmsg+"<br>"+"<li>请输入您的密码(长度不能大于20)。"
	founderr=true
else
	password=request.form("password")
end if
if password<>request("password2") then
	errmsg=errmsg+"<br>"+"<li>您输入的密码和确认密码不一致。"
	founderr=true
end if
if IsValidEmail(trim(request.form("Email")))=false then
	errmsg=errmsg+"<br>"+"<li>您的Email有错误。"
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
		errmsg="<br>"+"<li>对不起，您输入的用户名已经被注册，请重新输入。"
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
          <TD colSpan=2 height=24 bgcolor=#414141><b>会员注册成功</b></TD>
        </TR>
        <TR>
          <TD width=20% align=right>注 册 名</TD>
          <TD><%=user%> 　</TD>
        </TR>
        <TR>
          <TD align=right>性 别</TD>
          <TD><%if sex=1 then%>男<%else%>女<%end if%></TD>
        </TR>
        <TR>
          <TD align=right>密 码</TD>
          <TD><%=password%>　</TD>
        </TR>
        <TR>
          <TD align=right>Email地址</TD>
          <TD><%=email%>　</TD></TR>
        <TR>
          <TD align=right>OICQ号码</TD>
          <TD><%if request.form("oicq")="" then%>未注册<%else%><%=request.form("oicq")%><%end if%> </TD>
        </TR>
        <TR>
          <TD align=right>会员积分</TD>
          <TD>0</TD>
        </TR>
        <TR>
          <TD align=right>会员等级</TD>
          <TD>　</TD>
        </TR>
              <TR>
          <TD align=center colSpan=2><a href=index.asp>请登陆</a></TD>
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