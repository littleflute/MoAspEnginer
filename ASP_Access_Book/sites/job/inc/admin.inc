<% if session("flag")="" then
response.write"<SCRIPT language=JavaScript>alert('������ʱ�������µ�¼��');"
response.write"this.location.href='../';</SCRIPT>"
end if %>
<!--#include file="../inc/dbconn.inc"-->