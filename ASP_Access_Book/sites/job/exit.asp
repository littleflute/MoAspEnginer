<% if session("puid")<>"" then uname=session("puid") end if
   if session("cuid")<>"" then uname=session("cuid") end if
Session.Abandon
response.write"<SCRIPT language=JavaScript>alert('�û�"&uname&"�ɹ��˳���¼�����ڷ�����ҳ��');"
response.write"this.location.href='./';</SCRIPT>"  %>
