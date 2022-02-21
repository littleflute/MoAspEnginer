<!--#include file="conn.asp"-->
<!--#include file="isadmin.asp"-->
<%
sql = "select * from config"
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
<title>系统设置</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<br>
<form action="configok.asp" method="post">
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width="400" border="1" align="center" cellpadding=1>
<tr><td align="center" class="header" colspan=2>系统设置</td></tr>
  <tr>
    <td align="center">学校名称</td>
    <td>&nbsp;<input type=text name="schoolname" size=15 value=<%=rs("schoolname")%>></td>
  </tr>
  <tr>
    <td align="center">授权序列号</td>
    <td>&nbsp;<input type=text name="sn" size=15 value=<%=rs("sn")%>></td>
  </tr>
  <tr>
    <td align="center">管理员密码</td>
    <td>&nbsp;<input type=password name="adminpwd" size=15>（不填则保留原密码）</td>
  </tr>
  <tr>
    <td align="center">再输入一遍密码</td>
    <td>&nbsp;<input type=password name="adminpwd1" size=15></td>
  </tr>
  <tr>
    <td align="center">一级分类名</td>
    <td>&nbsp;<input type=text name="strfenlei1" size=15 value=<%=rs("strfenlei1")%>></td>
  </tr>
  <tr>
    <td align="center">二级分类名</td>
    <td>&nbsp;<input type=text name="strfenlei2" size=15 value=<%=rs("strfenlei2")%>></td>
  </tr>
  <tr>
    <td align="center">管理员E-Mail</td>
    <td>&nbsp;<input type=text name="adminmail" size=15 value=<%=rs("adminmail")%>></td>
  </tr>
  <tr>
    <td align="center">管理员联系方式</td>
    <td>&nbsp;<input type=text name="contactadmin" size=15 value=<%=rs("contactadmin")%>></td>
  </tr>
  <tr>
    <td align="center">系统URL</td>
    <td>&nbsp;<input type=text name="siteurl" size=15 value=<%=rs("siteurl")%>>（的网址）</td>
  </tr>
  <tr>
    <td align="center">每页显示数量</td>
    <td>&nbsp;<input type=text name="pagesize" size=15 value=<%=rs("pagesize")%>>资料列表的每页显示数量</td>
  </tr>
  <tr>
    <td align="center">是否锁定新用户</td>
    <td>
<%
if rs("locked") = 1 then
 response.write "&nbsp;<input type=radio name=locked value='0'>不锁定"
 response.write "<input type=radio name=locked value='1' checked>锁定"
else
 response.write "&nbsp;<input type=radio name=locked value='0' checked>不锁定"
 response.write "<input type=radio name=locked value='1'>锁定"
end if
%>
    </td>
  </tr>
  <tr>
    <td align="center">上传方式</td>
    <td>
        &nbsp;<select name="upload_type">
        <option value=0<%if rs("upload_type")=0 then response.write(" selected")%>>无组件上传</option>
        <option value=1<%if rs("upload_type")=1 then response.write(" selected")%>>Lyfupload</option>
        <option value=2<%if rs("upload_type")=2 then response.write(" selected")%>>Aspupload3.0</option>
        <option value=3<%if rs("upload_type")<>0 and rs("upload_type")<>1 and rs("upload_type")<>2 then response.write(" selected")%>>关闭上传功能</option>
        </select>
    </td>
  </tr>
  <tr>
    <td align="center">上传大小限制</td>
    <td>&nbsp;<input type=text name="upload_size" size=15 value=<%=rs("upload_size")%>>K</td>
  </tr>
  <tr>
    <td align="center">禁止上传的类型</td>
    <td>&nbsp;<input type=text name="upload_filetype" size=40 value=<%=rs("upload_filetype")%>><br>
        &nbsp;（扩展名之间请以逗号分隔）
    </td>
  </tr>
  <tr>
    <td align="center">管理员公告</td>
    <td>&nbsp;<textarea name="gonggao" cols="35" rows="4"><%=rs("gonggao")%></textarea></td>
  </tr>
  <tr>
    <td align="center">页面顶部<br>广告代码</td>
    <td>&nbsp;<textarea name="headads" cols="35" rows="4"><%=rs("headads")%></textarea></td>
  </tr>
</table>
<center><input type=submit name="submit" value="确认">&nbsp;&nbsp;<input type=reset name="reset" value="重填"></center>
</form>
</body>
</html>
<%
rs.close
set rs = nothing
conn.close
set conn = nothing
%>