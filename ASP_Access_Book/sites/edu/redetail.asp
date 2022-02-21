<html>
<head>
<title>回复作业</title>
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
</head>

<body bgcolor="#FFFFFF" text="#000000">
<div align="center">
<!--#include file="head.asp"-->
  <form action="redetailok.asp" method="post" >
    <table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=1 width=700 border=1>
      <tr>
        <td align="center" class="header" colspan=2>提交作业答案</td>
      </tr>
  <tr>
        <td width=164 align="center">学生姓名</td>
        <td width="520">&nbsp; 
          <input type=text name="name" size=20 >
          <input type=hidden name="reid" size=20 value=<%=request("id")%>>
        </td>
  </tr>
  <tr>
        <td align="center" width="164">作业标题</td>
        <td width="520">&nbsp; 
          <input type=text name="title" size=20>
        </td>
  </tr>

  <tr>
        <td align="center" width="164">作业答案</td>
        <td width="520"> 
          <textarea name="message" cols="40" rows="10"></textarea>
        </td>
  </tr>
<tr><td colspan=2 align=center>
          <input type=submit name="submit" value="提交">
          &nbsp;&nbsp;
<input type=reset name="reset" value="清空">
</td></tr>
</table>
</form></div>

<!--#include file="foot.asp"-->
</body>
</html>
