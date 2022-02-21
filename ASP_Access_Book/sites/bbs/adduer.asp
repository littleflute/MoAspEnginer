<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="an.htm" -->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid; background-color: #e3f1d1;">
  <tr> 
    <td width="475" height="24"><a href="index.asp">论坛论坛</a>→用户注册</td>
    <td width="283">|<a href="uerlist.asp">用户列表</a> | <a href="adduer.asp">用户注册</a>| 
      <a href="find.asp">查找会员</a>| <a href="uerlist.asp?order=2">发贴排名</a>|</td>
  </tr>
</table>
<table width="760" border="0" align="center">
  <tr > 
    <td height="22" colspan="2" background="pic/h3.gif"><div align="center"><font color="#FFFFFF"><strong>用户注册</strong></font></div></td>
  </tr>
  <tr> 
    <td colspan="2">
	<form ACTION="addsave.asp" METHOD="POST" name="form" id="form">
        <table width="760" border="1" cellpadding="0" cellspacing="1" bordercolor="#CCCCCC">
          <tr> 
            <td colspan="2" bgcolor="#e3f1d1">－－说明：此注册信息同时会加入协会通讯录，请认真如实填写</td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"> <div align="right">用户名*：</div></td>
            <td width="465"><input name="id" type="text" id="id"></td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"> <p align="right">密 码*：</td>
            <td width="465"> <input name="psw" type="password" id="psw"></td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"><div align="right">请再输入一次密码：</div></td>
            <td width="465"> <input name="pswc" type="password" id="psw0" size="20"></td>
          </tr>
          <tr> 
            <td width="238" height="18" bgcolor="#e3f1d1"> <div align="right">性 
                别*：</div></td>
            <td width="465"><input name="sex" type="radio" value="boy" checked>
              帅哥 &nbsp; <input type="radio" name="sex" value="girl">
              美女</td>
          </tr>
          <tr> 
            <td width="238" height="46" bgcolor="#e3f1d1"> <div align="right">个性化头像：<br>
              </div></td>
            <td width="465"><select                                  
      class=t2                                 
      onChange="document.images['idface'].src=options[selectedIndex].value;"                                 
      size=1 name=face>
                <option value=images/01.gif>用户头像-01 
                <option                
        value=images/02.gif>用户头像-02 
                <option value=images/03.gif>用户头像-03 
                <option                
        value=images/04.gif>用户头像-04 
                <option value=images/05.gif>用户头像-05 
                <option                
        value=images/06.gif>用户头像-06 
                <option value=images/07.gif>用户头像-07 
                <option                
        value=images/08.gif>用户头像-08 
                <option value=images/09.gif>用户头像-09 
                <option                
        value=images/10.gif>用户头像-10 
                <option value=images/11.gif>用户头像-11 
                <option                
        value=images/12.gif>用户头像-12 
                <option value=images/13.gif>用户头像-13 
                <option                
        value=images/14.gif>用户头像-14 
                <option value=images/15.gif>用户头像-15 
                <option                
        value=images/16.gif>用户头像-16 
                <option value=images/17.gif>用户头像-17 
                <option                
        value=images/18.gif>用户头像-18 
                <option value=images/19.gif>用户头像-19 
                <option                
        value=images/20.gif selected>用户头像-20</option>
              </select> <img src="images/20.gif" alt=个人形象代表 align="middle" class=t2 id=idface><font class=j1>[<a 
      href="tx.htm" target=_blank>所有头像</a>]</font></td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"><div align="right">Q Q：</div></td>
            <td width="465"><input name="qq" type="text" id="qq2"></td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"> <p align="right">E-mail*：</td>
            <td width="465"><input name="email" type="text" id="email2"> </td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"><div align="right">个人主页：</div></td>
            <td width="465"><input name="homepage" type="text" id="homepage2" value="http://"> 
            </td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"> <div align="right">*是否创协会员：</div></td>
            <td width="465"> <input name="huiyuan" type="radio" value="yes">
              是 &nbsp; <input type="radio" name="huiyuan" value="no">
              否</td>
          </tr>
          <tr> 
            <td colspan="2" bgcolor="#e3f1d1">－－以下将加入协会通信录，请务必认真填写</td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"> <div align="right">姓名*：</div></td>
            <td width="465"><input name="name" type="text" id="name"></td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"> <div align="right">部门：</div></td>
            <td width="465"> <select name="part" id="part">
                <option value="非会员">非会员</option>
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
            <td width="238" bgcolor="#e3f1d1"> <div align="right">职务：</div></td>
            <td width="465"> <select name="class" id="class">
                <option value="非会员">非会员</option>
                <option value="部长">部长</option>
                <option value="副部长">副部长</option>
                <option value="会员">会员</option>
                <option value="会长">会长</option>
                <option value="副会长">副会长</option>
              </select> </td>
          </tr>
          <tr> 
            <td bgcolor="#e3f1d1"><div align="right">院系&amp;年级：</div></td>
            <td><input name="college" type="text" id="college"></td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"> <div align="right">电 话：</div></td>
            <td width="465"> <input name="tel" type="text" id="tel"></td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"> <div align="right">住址：</div></td>
            <td width="465"> <input name="addr" type="text" id="addr"></td>
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
<!--#include file="foot.asp"-->