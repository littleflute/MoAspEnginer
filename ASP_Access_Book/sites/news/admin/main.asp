<%
  	if session("admin")="" then
  		response.redirect "admin.asp"
  	end if
%>
<html>
<link rel="stylesheet" href="style.css">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr>
    <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
        <tr> 
          <td background="../c.jpg" bgcolor="#FFFFFF"><img src="../c.jpg" width="536" height="54"></td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF"><div align="center"> 
              <table border="0" cellspacing="1" width="80%" align=center>
                <tr> 
                  <td width="100%"><br> <p align=center><b>信息发布管理系统</b></p>
                    <b>管理员可进行操作说明</b>：<br> <br>
                    1，通过Web添加文章，请选择添加软件。操作用户：所有用户 <br> <br>
                    2，对已经添加文章修改或删除，请点左边相关连接进行操作。操作用户：超级用户，普通管理员 <br> <br>
                    3，对栏目进行添加修改删除，请点左边相关连接进行操作。操作用户：超级用户 <br> <br>
                    4，对地区类别进行添加修改删除，请点左边相关连接进行操作。操作用户：超级用户 <br> <br>
                    5，用户的增加修改删除，请点左边相关连接进行操作。操作用户：超级用户 <br> <br>
                    6，<font color=red>为了系统的安全性，离开管理请点击退出系统</font> <br> </td>
                </tr>
              </table>
            </div></td>
        </tr>
        <tr> 
          <td background="../b.jpg" bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</html>