<% If session("AdminUID")="" Then %><table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid;">
  <tr> 
    <td background="pic/h3.gif"><strong>-=&gt; <font color="#FFFFFF">快速登录入口</font></strong><font color="#0000FF">&nbsp; 
      </font>[<a href="adduer.asp">注册用户</a>]　</td>
  </tr>
  <form method="POST" action="loginCheck.asp">
    <tr> 
      <td height="21" bgcolor="#f4faed"> 用户名： 
        <input type="text" name="AdminUID" size="20">
        密码： 
        <input type="password" name="AdminPWD" size="20"> <input type="submit" value="提交" name="B1" style="font-family: 宋体; font-size: 9pt; background-color: #FFFFFF; border-style: solid; border-width: 1"> 
      </td>
    </tr>
  </form>
</table>
<% End If %>
