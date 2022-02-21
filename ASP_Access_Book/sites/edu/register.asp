<!--#include file="conn.asp"-->
<!--#include file="fenlei.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE>TD {
	FONT-SIZE: 9pt; LINE-HEIGHT: 140%
}
BODY {
	FONT-SIZE: 9pt; LINE-HEIGHT: 140%
}
A:link {
	COLOR: #0033cc; TEXT-DECORATION: none
}
A:visited {
	COLOR: #0033cc; TEXT-DECORATION: none
}
A:active {
	COLOR: #ff0000; TEXT-DECORATION: none
}
A:hover {
	COLOR: #000000; TEXT-DECORATION: underline
}
.header	{
font-family: Tahoma, Verdana; font-size: 9pt; color: #FFFFFF; background-color: #00CCCC
}
.category{
font-family: Tahoma, Verdana; font-size: 9pt; color: #000000; background-color: #EFEFEF
}
</STYLE>
<title>教师注册</title>
<script language="javascript">
function submitonce(theform)
{
//if IE 4+ or NS 6+
if (document.all||document.getElementById){
//screen thru every element in the form, and hunt down "submit" and "reset"
for (i=0;i<theform.length;i++){
var tempobj=theform.elements[i]
if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
//disable em
tempobj.disabled=true
}
}
}
</script>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<!--#include file="head.asp"-->
<form action="registerok.asp" method="post" onSubmit=submitonce(this)>
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width="400" border="1" align="center" cellpadding=1>
<tr><td align="center" class="header" colspan=2>教师注册</td></tr>
  <tr>
    <td width=130>&nbsp;请输入您所属的<%=strfenlei1%></td>
    <td width=270>&nbsp;<input type=text name="fenlei1" size=25></td>
  </tr>
  <tr>
    <td>&nbsp;请输入您所属的<%=strfenlei2%></td>
    <td>&nbsp;<input type=text name="fenlei2" size=25></td>
  </tr>
  <tr>
    <td>&nbsp;您的姓名</td>
    <td>&nbsp;<input type=text name="teacher" size=25></td>
  </tr>
  <tr>
    <td>&nbsp;您的登录名</td>
    <td>&nbsp;<input type=text name="loginname" size=22>（不能全为数字）</td>
  </tr>
  <tr>
    <td>&nbsp;请输入密码</td>
    <td>&nbsp;<input type=password name="password" size=25></td>
  </tr>
  <tr>
    <td>&nbsp;再输一遍密码</td>
    <td>&nbsp;<input type=password name="password1" size=25></td>
  </tr>
  <tr>
    <td>&nbsp;密码找回问题</td>
    <td>&nbsp;<input type=text name="ask" size=25></td>
  </tr>
  <tr>
    <td>&nbsp;密码找回答案</td>
    <td>&nbsp;<input type=text name="answer" size=25></td>
  </tr>
  <tr>
    <td>&nbsp;E-Mail</td>
    <td>&nbsp;<input type=text name="email" size=25>（可以不填）</td>
  </tr>
  <tr>
    <td>&nbsp;个人主页</td>
    <td>&nbsp;<input type=text name="homepage" size=25>（可以不填）</td>
  </tr>
  <tr>
    <td>&nbsp;QQ号码</td>
    <td>&nbsp;<input type=text name="qq" size=25>（可以不填）</td>
  </tr>
  <tr>
    <td>&nbsp;通讯地址</td>
    <td>&nbsp;<input type=text name="address" size=25>（可以不填）</td>
  </tr>
  <tr>
    <td>&nbsp;个人简介</td>
    <td>
        &nbsp;<textarea name="intro" cols="25" rows="6"></textarea>
    </td>
  </tr>
</table>
<center><br><input type=submit name="submit" value="确认">&nbsp;&nbsp;<input type=reset name="reset" value="清空"></center>
</form>
<!--#include file="foot.asp"-->
</body>
</html>
<%
conn.close
set conn = nothing
%>