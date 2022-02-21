<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/md5code.asp" -->
<!--#include file="include/library.asp" -->
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
    <td width="128" align="center" valign="top"><br>
      <br>
      <strong>用户注册</strong></td>
    <td><br><%IF Request.QueryString("action")="agree" Then%>
      <table width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
        <form name="register" method="post" action="register.asp?action=register"><tr class="msg_head">
          <td height="24" colspan="2" align="center"><strong>用户注册</strong></td>
        </tr>
          <tr>
            <td align="right" bgcolor="#FFFFFF">用户名： </td>
            <td bgcolor="#FFFFFF"><input name="UserName" type="text" id="UserName" size="24" maxlength="24">
            <b><font color="#FF0000">&nbsp;*</font></b> 用户名不能超过24个字符，12个汉字</td>
          </tr>
          <tr>
            <td align="right" bgcolor="#FFFFFF">密码： </td>
            <td bgcolor="#FFFFFF"><input name="Password" type="password" id="Password" size="16" maxlength="16">
            &nbsp;<b><font color="#FF0000">*</font></b> 密码必须是6-16位</td>
          </tr>
          <tr>
            <td align="right" bgcolor="#FFFFFF">确认密码：</td>
            <td bgcolor="#FFFFFF"><input name="PasswordR" type="password" id="PasswordR" size="16" maxlength="16">
            &nbsp;<b><font color="#FF0000">*</font></b> 请输入确认密码</td>
          </tr>
          <tr>
            <td align="right" bgcolor="#FFFFFF">电子信箱：</td>
            <td bgcolor="#FFFFFF"><input name="Email" type="text" id="Email">
            &nbsp;<b><font color="#FF0000">*</font></b> 请输入有效的电子信箱地址</td>
          </tr>
          <tr>
            <td align="right" bgcolor="#FFFFFF">验证码：</td>
            <td bgcolor="#FFFFFF"><img src="include/validatecode.asp" align="absmiddle" border="0"> <input name="validatecode" type="text" id="validatecode" size="4">
            &nbsp;<b><font color="#FF0000">*</font></b> 为了防止恶意注册，请输入验证码</td>
          </tr>
        <tr>
          <td colspan="2" align="center" bgcolor="#FFFFFF"><input name="Submit" type="submit" id="Submit" value=" 提 交 ">
&nbsp;<input name="Reset" type="reset" id="Reset" value=" 重 置 ">
          </td>
        </tr></form>
      </table>
      <%ElseIF Request.QueryString("action")="register" Then
	  Dim Reg_Title,Reg_Message
	  IF CheckStr(Request.Form("UserName"))=Empty OR CheckStr(Request.Form("Password"))=Empty OR CheckStr(Request.Form("PasswordR"))=Empty OR CheckStr(Request.Form("Email"))=Empty Then
	  	Reg_Title="错误信息"
		Reg_Message="请将信息输入完整<br><a href='javascript:history.go(-1);'>请返回重新输入</a>"
	Else
		IF  CheckStr(Request.Form("Password"))<>CheckStr(Request.Form("PasswordR")) Then
			Reg_Title="错误信息"
			Reg_Message="密码与确认密码不符<br><a href='javascript:history.go(-1);'>请返回重新输入</a>"
		ElseIF Len(CheckStr(Request.Form("Password")))<6 OR Len(CheckStr(Request.Form("Password")))>16 Then
			Reg_Title="错误信息"
			Reg_Message="密码长度不符合<br><a href='javascript:history.go(-1);'>请返回重新输入</a>"
		ElseIF Len(CheckStr(Request.Form("UserName")))>24 Then
			Reg_Title="错误信息"
			Reg_Message="用户名长度超过24个字符，12个汉字<br><a href='javascript:history.go(-1);'>请返回重新输入</a>"
		ElseIF IsValidUserName(Request.Form("UserName"))=False Then
			Reg_Title="错误信息"
			Reg_Message="用户名中含有非法字符<br><a href=""javascript:history.go(-1);"">请返回重新输入</a>"
		ElseIF IsValidEmail(CheckStr(Request.Form("Email")))=False Then
			Reg_Title="错误信息"
			Reg_Message="电子信箱输入错误<br><a href=""javascript:history.go(-1);"">请返回重新输入</a>"
		ElseIf Trim(Session("L-Blog_ValidateCode"))<>Trim(Request.Form("validatecode")) Then
			Reg_Title="错误信息"
			Reg_Message="验证码输入错误！<br><a href=""javascript:history.go(-1);"">请返回重新输入</a>"
		Else
			Dim memAlready
			Set memAlready=Conn.Execute("SELECT mem_Name FROM blog_Member WHERE mem_Name='"&CheckStr(Request.Form("Username"))&"'")
			SQLQueryNums=SQLQueryNums+1
			IF NOT(memAlready.EOF AND memAlready.BOF) Then
				Reg_Title="错误信息"
				Reg_Message="对不起，你所使用的用户名已经注册<br><a href='javascript:history.go(-1);'>请返回重新输入</a>"
			Else
				Conn.Execute("INSERT INTO blog_member(mem_Name,mem_Password,mem_Email,mem_LastIP) Values ('"&CheckStr(Request.Form("Username"))&"','"&md5(CheckStr(Request.Form("Password")))&"','"&CheckStr(Request.Form("Email"))&"','"&Guest_IP&"')")
				Conn.ExeCute("UPDATE blog_Info SET blog_MemNums=blog_MemNums+1")
				SQLQueryNums=SQLQueryNums+2
				Response.Cookies(CookieName)("memName")=CheckStr(Request.Form("Username"))
				Response.Cookies(CookieName)("memPassword")=md5(CheckStr(Request.Form("Password")))
				Response.Cookies(CookieName)("memStatus")="Member"
				Reg_Title="注册成功"
				Reg_Message="欢迎你来到 "&siteName&"<br><a href='default.asp'>点击返回主页</a>"
			End IF
			Set memAlready=Nothing
		End IF
	End IF%>
      <table width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
          <tr>
            <td height="24" align="center"><strong><%=Reg_Title%></strong></td>
          </tr>
          <tr>
            <td height="88" align="center" valign="middle" bgcolor="#FFFFFF"><%=Reg_Message%></td>
          </tr>
      </table>
      <%Else%><table width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
      <tr class="msg_head">
        <td height="24" align="center"><strong>用户注册协议</strong></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF">网站测试：暂无内容<br>
　                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         </td>
      </tr>
      <tr>
        <form name="bbrules" method="post" action="register.asp?action=agree"><td align="center" bgcolor="#FFFFFF"><input name="rulesubmit" type="submit" id="agree" value="我已阅读并同意以上条款">
&nbsp;    <input type="button" name="return" value="不同意" onClick="javascript:history.go(-1);">
	<script language="javascript">
	var secs = 9;
	var wait = secs * 1000;
	document.bbrules.rulesubmit.value = "我已阅读并同意以上条款 (" + secs + ")";
	document.bbrules.rulesubmit.disabled = true;
	for(i = 1; i <= secs; i++) {
	        window.setTimeout("update(" + i + ")", i * 1000);
	}
	window.setTimeout("timer()", wait);
	function update(num, value) {
	        if(num == (wait/1000)) {
	                document.bbrules.rulesubmit.value = "我已阅读并同意以上条款";
	        } else {
	                printnr = (wait / 1000)-num;
	                document.bbrules.rulesubmit.value = "我已阅读并同意以上条款 (" + printnr + ")";
	        }
	}
	function timer() {
	        document.bbrules.rulesubmit.disabled = false;
	        document.bbrules.rulesubmit.value = "我已阅读并同意以上条款";
	}
	</script></td></form>
      </tr>
    </table><%End IF%>
    <br></td>
  </tr>
</table>
<!--#include file="footer.asp" -->