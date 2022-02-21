<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>洋洋购物电子商务网站资料上传管理界面</title>
</head>

<body>

<%
	if Request("T1") <> "Homestar" or Request("T2") <> "hd51988" then
%>
<div align="center">
  <center><form name="CheckAccount" action="upld.asp" method="POST">
  <table border="0" cellpadding="0" cellspacing="0" width="30%">
    <tr>
      <td width="100%" colspan="2">
        <p align="center"><font size="2"><b>管理员登录:</b></font></p>
      </td>
    </tr>
    <tr>
      <td width="31%" align="right"><b><font size="2">账号</font></b></td>
      <td width="69%"><input type="text" name="T1" size="20"></td>
    </tr>
    <tr>
      <td width="31%" align="right"><b><font size="2">密码</font></b></td>
      <td width="69%"><input type="password" name="T2" size="20"></td>
    </tr>
    <tr>
      <td width="31%"></td>
      <td width="69%"><input type="submit" value="登录" name="B1">&nbsp; <input type="reset" value="重写" name="B2"></td>
    </tr>
  </table></form>
  </center>
</div>

<%
	else
%>

<form name="me">
  <p align="center"><input type="button" value="新产品内容的上传" name="B1" onClick="
  			me.action='sheet.asp?OP=新增'
  			me.method='post'
  			me.submit()
  			"></p>
  <p align="center"><input type="button" value="管理重设产品标记" name="B2" onClick="
  			me.action='tag_list.asp'
  			me.method='post'
  			me.submit()
  			"></p>
  <p align="center"><input type="button" value="管理所有产品内容" name="B3" onClick="
  			me.action='pro_list.asp'
  			me.method='post'
  			me.submit()
  			"></p>
</form>

<%
	end if
%>
</body>

</html>
