<% if session("puid")<>"" then uname=session("puid") end if
   if session("cuid")<>"" then uname=session("cuid") end if
Session.Abandon
response.write"<SCRIPT language=JavaScript>alert('用户"&uname&"成功退出登录，正在返回首页！');"
response.write"this.location.href='./';</SCRIPT>"  %>
