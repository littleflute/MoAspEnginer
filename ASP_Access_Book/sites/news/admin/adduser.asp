<%@language=vbscript%>
<%
  if session("admin")="" then
  response.redirect "admin.asp"
  else
	if session("flag")>1 then
		response.write "<br><p align=center>您没有操作的权限</p>"
		response.end
	end if
  end if
  
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<script language=javascript>

function check()

{
  if(document.form1.username.value=="")
    {
      alert("用户名为空");
      return false;
    }
    
  if(document.form1.newpin.value=="")
    {
      alert("密码不能为空");
      return false;
    }
    
  if((document.form1.newpin.value)!=(document.form1.re_newpin.value))
    {
      alert("密码不匹配");
      return false;
    }

}

</script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<div align="center">
  <p>&nbsp;</p>
  <form method="post" action="saveuser1.asp" name="form1" onsubmit="javascript:return check();">
    <table width="40%" border="1" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
      <tr bgcolor="#CCCCCC"> 
        <td height="25"> 
          <div align="center"><font size="2">新增管理员</font></div>
        </td>
      </tr>
      <tr> 
        <td height="30"> 
          <div align="center"><font size="2">用 户 名　 
            <input type="text" name="username">
            </font></div>
        </td>
      </tr>
      <tr> 
        <td height="30"> 
          <div align="center"><font size="2">初始密码　 
            <input type="password" name="newpin">
            </font> </div>
        </td>
      </tr>
      <tr> 
        <td height="30"> 
          <div align="center"><font size="2">确认密码　 
            <input type="password" name="re_newpin">
            </font></div>
        </td>
      </tr>
      <tr> 
        <td height="30"> 
          <div align="center">权限设置 
             </div>
        </td>
      </tr>
      <tr>
        <td height="30">
          <div align="center">
            <input type="radio" name="right_class" value="3">
            用户
            <input type="radio" name="right_class" value="2" checked>
            普通管理员
            <input type="radio" name="right_class" value="1">
            超级用户</div>
        </td>
      </tr>
    </table>
    <p>
      <input type="submit" name="Submit" value="确定">
    </p>
  </form>
  <p>&nbsp;</p>
  <p align=center></p> </div>
</body>
</html>
