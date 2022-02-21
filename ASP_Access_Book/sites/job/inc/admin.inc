<% if session("flag")="" then
response.write"<SCRIPT language=JavaScript>alert('操作超时，请重新登录！');"
response.write"this.location.href='../';</SCRIPT>"
end if %>
<!--#include file="../inc/dbconn.inc"-->