<!--#include file="conn.asp"-->
<%
if session("admin")="" then
	response.redirect "admin.asp"
	response.end
end if
%>

<%
set rs=server.createobject("adodb.recordset")
sql="select * from picclass"
rs.open sql,conn,1,1
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>添加图片</title>
<link rel=stylesheet href="style.css" type="text/css">
</head>

<body>

<div align="center">
  <center>
  <table border="0" width="400" cellspacing="0" cellpadding="0">
    <tr>
      <td width="100%">
        <form method="POST" action="add_pic1.asp">
          <table border="0" width="100%" cellspacing="0" cellpadding="3">
            <tr>
              <td width="100%" align="right" colspan="2">
                <p align="center">添加产品图片</td>
            </tr>
            <tr>
              <td width="37%" align="right">产品名称：</td>
              <td width="63%"><input type="text" name="pic_name" size="20"></td>
            </tr>
            <tr>
              <td width="37%" align="right">产品型号：</td>
              <td width="63%"><input type="text" name="pic_type" size="20"></td>
            </tr>
            <tr>
              <td width="37%" align="right">产品类型：</td>
              <td width="63%"><select size="1" name="pic_class">
              <%
              do while not rs.eof
              %>
                  <option value=<%=rs("pic_class")%>><%=rs("pic_class")%></option>
                  <%rs.movenext
                  loop
                  rs.close
                  set rs=nothing
                  %>
                </select></td>
            </tr>
            <tr>
              <td width="37%" align="right">产品说明：</td>
              <td width="63%"><textarea rows="2" name="pic_text" cols="20"></textarea></td>
            </tr>
            <tr>
              <td width="37%" align="right">小图片网址：</td>
              <td width="63%"><input type="text" name="pic_lurl" size="20" value="http://"></td>
            </tr>
            <tr>
              <td width="37%" align="right">大图片网址：</td>
              <td width="63%"><input type="text" name="pic_url" size="20" value="http://"></td>
            </tr>
            <tr>
              <td width="100%" align="right" colspan="2">
                <p align="right"><input type="submit" value="提交" name="B1"><input type="reset" value="全部重写" name="B2"></td>
            </tr>
          </center>
        </table>
        <center>
        <p>　</p>
        </form>
        <p>　</td>
    </tr>
  </table>
</div>

</body>

</html>
