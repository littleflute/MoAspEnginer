<% uname=session("flag")
Session.Abandon
response.write"<SCRIPT language=JavaScript>alert('�û�"&uname&"�ɹ��˳���¼��');"
response.write"this.location.href='../';</SCRIPT>"  %>
