<% If session("AdminUID")="" Then %><table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid;">
  <tr> 
    <td background="pic/h3.gif"><strong>-=&gt; <font color="#FFFFFF">���ٵ�¼���</font></strong><font color="#0000FF">&nbsp; 
      </font>[<a href="adduer.asp">ע���û�</a>]��</td>
  </tr>
  <form method="POST" action="loginCheck.asp">
    <tr> 
      <td height="21" bgcolor="#f4faed"> �û����� 
        <input type="text" name="AdminUID" size="20">
        ���룺 
        <input type="password" name="AdminPWD" size="20"> <input type="submit" value="�ύ" name="B1" style="font-family: ����; font-size: 9pt; background-color: #FFFFFF; border-style: solid; border-width: 1"> 
      </td>
    </tr>
  </form>
</table>
<% End If %>
