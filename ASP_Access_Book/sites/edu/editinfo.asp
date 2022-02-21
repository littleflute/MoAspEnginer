<!--#include file="conn.asp"-->
<!--#include file="fenlei.asp"-->
<!--#include file="isteacher.asp"-->
<%
id = request("id")
if id = ""  then
 if session("teacherid")<>request("id") then
 conn.close

 set conn = nothing
 response.write "<script>alert('请不要捣乱!!!');top.window.location.href='teachermain.asp';</script>"
 response.end
  end if
end if
sql = "select * from teacher where teacherid="&id
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
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
<title>修改资料</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<br>
<form action="editinfook.asp" method="post">
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width="480" border="1" align="center" cellpadding=1>
<tr><td align="center" class="header" colspan=2>修改资料</td></tr>
  <tr>
    <td width=100>所属<%=strfenlei1%></td>
    <td width=380><input type=text name="fenlei1" size=25 value="<%=rs("fenlei1")%>"></td>
  </tr>
  <tr>
    <td>所属<%=strfenlei2%></td>
    <td><input type=text name="fenlei2" size=25 value="<%=rs("fenlei2")%>"></td>
  </tr>
  <tr>
    <td>您的姓名</td>
    <td><input type=text name="teacher" size=25 value="<%=rs("teacher")%>"></td>
  </tr>
  <tr>
    <td>您的登录名</td>
    <td><input type=text name="loginname" size=25 value="<%=rs("loginname")%>">（不能全为数字）</td>
  </tr>
  <tr>
    <td>密码</td>
    <td><input type=password name="password" size=25>（不填则保留原密码）</td>
  </tr>
  <tr>
    <td>重复密码</td>
    <td><input type=password name="password1" size=25></td>
  </tr>
  <tr>
    <td>密码找回问题</td>
    <td><input type=text name="ask" size=25 value="<%=rs("ask")%>"></td>
  </tr>
  <tr>
    <td>密码找回答案</td>
    <td><input type=text name="answer" size=25 value="<%=rs("answer")%>"></td>
  </tr>
  <tr>
    <td>E-Mail</td>
    <td><input type=text name="email" size=25 value="<%=rs("email")%>">（可以不填）</td>
  </tr>
  <tr>
    <td>个人主页</td>
    <td><input type=text name="homepage" size=25 value="<%=rs("homepage")%>">（可以不填）</td>
  </tr>
  <tr>
    <td>QQ号码</td>
    <td><input type=text name="qq" size=25 value="<%=rs("qq")%>">（可以不填）</td>
  </tr>
  <tr>
    <td>通讯地址</td>
    <td><input type=text name="address" size=25 value="<%=rs("address")%>">（可以不填）</td>
  </tr>
  <tr>
    <td align="rcenter">上传照片</td>
    <td>
        <iframe name="up" frameborder=0 width=380 height=26 scrolling=no src=upload.asp></iframe></td>
  </tr>
  <tr>
    <td>照片地址</td>
    <td><input type=text size=25 name="fileurl" value="<%=rs("photourl")%>">（可以不填）</td>
  </tr>
  <tr>
    <td>个人简介</td>
    <td>
        <textarea name="intro" cols="25" rows="6"><%=rs("intro")%></textarea>
    </td>
  </tr>
</table>
<center><input type=submit name="submit" value="更新"></center>
</form>
<%
rs.close
set rs = nothing
conn.close
set conn = nothing
%>
</body>
</html>