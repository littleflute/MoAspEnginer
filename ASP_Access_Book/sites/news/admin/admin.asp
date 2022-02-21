<html>

<head>
<title>管理者登陆</title>
<link rel="stylesheet" href="style.CSS">
</head>

<body>
<div align="center">
  <p>&nbsp;</p>
  <form method="post" action="Chkadmin.asp">
    <table class="border" width="300" border="0" cellpadding="4" cellspacing="0" >
      <tr class="title"> 
        <td colspan="2"> 
          <div align="center">操作: 确认身份</div>
        </td>
      </tr>
      <tr> 
        <td class="tdbg" colspan="2"> <br>
          <br>
          <table width="250" border="0" cellspacing="0" cellpadding="0" align="center">
            <tr> 
              <td>用户名称：
                <input class="smallinput"  type="text" name="Username" size="23">
                <br>
                用户密码：
                <input class="smallinput"  type="password" name="Password" size="23">
                <br>
                <br>
                <br>
              </td>
            </tr>
            <tr> 
              <td> 
                <div align="center"> 
                  <input class="buttonface"  type="submit" name="Submit" value="确认">
                  <input class="buttonface" type="reset" name="Submit2" value="复位">
                  <br>
                  <br>
                </div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </form>
  <p>&nbsp;</p>
</div>
</body>
</html>
