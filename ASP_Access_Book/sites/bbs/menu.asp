<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px #a9d46d solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;">
  <tr> 
    <td height="10" valign="top" background="pic/h1.gif" class="p9"><img src="" alt="" name="" width="32" height="1"></td>
  </tr>
  <tr > 
    <td height="16" background="pic/h2.gif" class="p9" > 
      |  <% If session("AdminUID")<>"" Then %>
      <%= session("AdminUID") %> ��ã� | <a href="useredit.asp?id=<%=session("AdminUID")%>">�޸�����</a>| 
      <a href=javascript:openScript('duan.asp',420,320)><FONT COLOR=#000000>����Ϣ</FONT></a>     <% If session("Adminflag")="0" Then %>
   | <a href="oneedit.asp">������̳</a> <% End If %>
      <% Else %>
      <a href="login.asp">��½</a> 
      <% End If %>
      | <a href="adduer.asp">�û�ע��</a> | <a href="uerlist.asp">�û��б�</a>| <a href="logout.asp">�˳���½</a></td>
  </tr>
</table>
