<% if session("cuid")<>"" and session("puid")<>"" then
   session.Abandon
   response.write"<SCRIPT language=JavaScript>alert('用户非法登录，正在返回首页，请重新登录！');"
   response.write"this.location.href='../';</SCRIPT>" 
   end if
   if session("cuid")="" then
   response.write"<SCRIPT language=JavaScript>alert('操作超时，请重新登录！');"
   response.write"this.location.href='../';</SCRIPT>" 
   end if 
%>
<!--#include file="../inc/dbconn.inc"-->