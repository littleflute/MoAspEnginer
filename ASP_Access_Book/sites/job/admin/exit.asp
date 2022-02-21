<% uname=session("flag")
Session.Abandon
response.write"<SCRIPT language=JavaScript>alert('用户"&uname&"成功退出登录！');"
response.write"this.location.href='../';</SCRIPT>"  %>
