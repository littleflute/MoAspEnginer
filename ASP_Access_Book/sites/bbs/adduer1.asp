<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!-- #include file="an.htm" -->
<table width="760" border="0" align="center">
  <tr> 
    <td colspan="2" bgcolor="#FFFF99"><div align="center">用户注册</div></td>
  </tr>
  <tr> 
    <td colspan="2">
	<form ACTION="addsave.asp" METHOD="POST" name="form" id="form">
        <table width="95%" border="1" cellpadding="0" cellspacing="1" bordercolor="#CCCCCC">
          <tr> 
            <td width="24%" bgcolor="#CCCCCC"> <div align="right">用户名：</div></td>
            <td width="76%"><input name="id" type="text" id="id"></td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"><div align="right">密 码：</div></td>
            <td> <input name="psw" type="password" id="psw"></td>
          </tr>
          <tr> 
            <td height="18" bgcolor="#CCCCCC"> <div align="right">性 别：</div></td>
            <td><input name="sex" type="radio" value="boy" checked>
              帅哥 &nbsp; <input type="radio" name="sex" value="girl">
              美女</td>
          </tr>
          <tr> 
            <td height="19" bgcolor="#CCCCCC"> <div align="right">Q Q：</div></td>
            <td><input name="qq" type="text" id="qq"></td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"><div align="right">E-mail：</div></td>
            <td><input name="email" type="text" id="email"></td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"><div align="right">个人主页：</div></td>
            <td><input name="homepage" type="text" id="homepage"> </td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"> <div align="right">是否创协会员：</div></td>
            <td> <input name="huiyuan" type="radio" value="yes" checked>
              是 &nbsp; <input type="radio" name="huiyuan" value="no">
              否</td>
          </tr>
          <tr> 
            <td colspan="2" bgcolor="#CCCCCC">以下将加入协会通信录，请务必认真填写</td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"> <div align="right">姓名：</div></td>
            <td bgcolor="#CCCCCC"><input name="name" type="text" id="name"></td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"> <div align="right">部门：</div></td>
            <td> <select name="part" id="part">
                <option value="会长">会长</option>
                <option value="副会长">副会长</option>
                <option value="秘书处">秘书处</option>
                <option value="项目部">项目部</option>
                <option value="公关部">公关部</option>
                <option value="宣传部">宣传部</option>
                <option value="OEC活动部">OEC活动部</option>
                <option value="会员俱乐部">会员俱乐部</option>
              </select></td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"> <div align="right">职务：</div></td>
            <td> <select name="class" id="class">
                <option value="会长">会长</option>
                <option value="副会长">副会长</option>
                <option value="部长">部长</option>
                <option value="副部长">副部长</option>
                <option value="会员">会员</option>
              </select> </td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"> <div align="right">院系&amp;年级：</div></td>
            <td> <input name="college" type="text" id="college"></td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"> <div align="right">电 话：</div></td>
            <td> <input name="tel" type="text" id="tel"></td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"> <div align="right">住址：</div></td>
            <td> <input name="addr" type="text" id="addr"></td>
          </tr>
        </table>
        <p align="center"> 
          <input type="submit" name="Submit" value="  添好了  ">
          <input type="reset" name="Submit2" value="  重 来  ">
        </p>
        <input type="hidden" name="MM_insert" value="form1">
      </form></td>
  </tr>
  <tr> 
    <td width="242">&nbsp;</td>
    <td width="508">&nbsp;</td>
  </tr>
</table>
</body>
</html>
