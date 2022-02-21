<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/md5code.asp" -->
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
    <td width="128" align="center" valign="top"><br>
    <br>
    <strong>用户登录</strong></td>
    <td align="center"><br>
	<%IF Request.QueryString("action")="logout" Then
	Response.Cookies(CookieName)("memName")=""
	Response.Cookies(CookieName)("memPassword")=""
	Response.Cookies(CookieName)("memStatus")=""%><div class="msg_head">退出成功</div>
          <div class="msg_content"><a href="default.asp">点击返回主页</a></div>
      <%ElseIF Request.QueryString("action")="login" Then
	  Dim Login_Title,Login_Message
	  IF Request.Form("UserName")=Empty OR Request.Form("Password")=Empty Then
	  	  Login_Title="错误信息"
		  Login_Message="请将信息输入完整<br><a href='javascript:history.go(-1);'>请返回重新输入</a>"
	 Else
		 	If Trim(Request.Form("validatecode"))=Empty Or Trim(Session("L-Blog_ValidateCode"))<>Trim(Request.Form("validatecode")) THEN
		 Login_Title="错误信息"
		 Login_Message="请输入<span style=""font-size:11px;color: Green;"">验证码</span>！"
		 Else
		 Dim memLogin
		 Set memLogin=Server.CreateObject("ADODB.Recordset")
		 SQL="SELECT mem_Name,mem_Password,mem_Status,mem_LastIP FROM blog_member WHERE mem_Name='"&CheckStr(Request.Form("Username"))&"' AND mem_Password='"&md5(CheckStr(Request.Form("Password")))&"'"
		 memLogin.Open SQL,Conn,1,3
		 SQLQueryNums=SQLQueryNums+1
		 IF memLogin.EOF AND memLogin.BOF Then
		 	Login_Title="错误信息"
		  	Login_Message="用户名与密码错误<br><a href='javascript:history.go(-1);'>请返回重新输入</a>"
		 Else
		 	Response.Cookies(CookieName)("memName")=memLogin("mem_Name")
			Response.Cookies(CookieName)("memPassword")=memLogin("mem_Password")
			Response.Cookies(CookieName)("memStatus")=memLogin("mem_Status")
			Select Case Request.Form("CookieTime")
				Case 1
					Response.Cookies(CookieName).Expires=Date+1
				Case 2
					Response.Cookies(CookieName).Expires=Date+31
				Case 3
					Response.Cookies(CookieName).Expires=Date+365
			End Select
			memLogin("mem_LastIP")=Guest_IP
			memLogin.Update
		 	Login_Title="登陆成功"
		  	Login_Message="<a href="""&Request.ServerVariables("HTTP_ReFeReR")&""">点击返回正在访问的页面</a><br><br><a href=""forumview.asp"">点击返回论坛首页</a><br><br><a href='default.asp'>点击返回主页</a>"
		 End IF
		  End IF
		 memLogin.Close
		 Set memLogin=Nothing
	End IF%>
      <div class="msg_head"><%=Login_Title%></div>
          <div class="msg_content"><%=Login_Message%></div>
      <%Else%><table width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC"><form name="Login" method="post" action="logging.asp?action=login">
      <tr>
        <td height="24" colspan="2" align="center" class="msg_head">用户登陆</td>
      </tr>
      <tr>
        <td align="right" bgcolor="#FFFFFF">用户：</td>
        <td bgcolor="#FFFFFF">          <input name="username" type="text" id="username" size="24" maxlength="24">        </td>
      </tr>
      <tr>
        <td align="right" bgcolor="#FFFFFF">密码：</td>
        <td bgcolor="#FFFFFF"><input name="password" type="password" id="password" size="16" maxlength="16">
          &nbsp;</td>
      </tr>
     <tr>
        <td align="right" bgcolor="#FFFFFF">验证码：</td>
        <td bgcolor="#FFFFFF"><input name="validatecode" type="text" id="validatecode" class="text" size="3" />&nbsp;<img src="include/validatecode.asp" align="absmiddle" border="0" />
          &nbsp;</td>
      </tr>
      <tr>
        <td align="right" bgcolor="#FFFFFF"> Cookie选项：<br></td>
        <td bgcolor="#FFFFFF"><input type="radio" name="CookieTime" value="0" checked>
          不保存
            <input type="radio" name="CookieTime" value="1">
            保存一天
            <input type="radio" name="CookieTime" value="2">
            保存一月
            <input type="radio" name="CookieTime" value="3">
            保存一年</td>
      </tr>
      <tr>
        <td colspan="2" align="center" bgcolor="#FFFFFF"><input name="Login" type="submit" id="agree" value=" 登 陆 ">
&nbsp;    <input name="Register" type="button" id="Register" onclick="javascript:document.location.href='register.asp';" value=" 注 册 ">
	</td>
      </tr></form>
    </table>
      <%End IF%>
    <br></td>
  </tr>
</table>
<!--#include file="footer.asp" -->