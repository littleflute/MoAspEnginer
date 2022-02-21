<!--#include file="conn.asp"-->
<%
if session("admin")="" then
	response.redirect "admin.asp"
	response.end
end if
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>管理产品类</title>
<link rel=stylesheet href="style.css" type="text/css">
</head>

<body>


<div align="center">
  <center>
  <table border="0" width="400" cellspacing="0" cellpadding="3">
      <%
    set rs=server.createobject("adodb.recordset")
    sql="select * from picclass"
    rs.open sql,conn,1,1
    do while not rs.eof
    %>
    <tr>
    
      <td width="219"><%=rs("pic_class")%></td>
      <td width="165"><a href=del_class.asp?id=<%=rs("id")%>>删除</a></td>
    </tr>
    <%
    rs.movenext
    loop
    rs.close
    set rs=nothing
    %>
  </table>
  </center>
</div>


<div align="center">
  <center>
  <table border="0" width="400" cellspacing="0" cellpadding="0">
    <tr>
      <td width="100%">
        <form method="POST" action="add_class.asp">
          <table border="0" width="100%" cellspacing="0" cellpadding="4">
            <tr>
              <td width="100%" colspan="2"></td>
            </tr>
            <tr>
              <td width="100%" colspan="2">
                <p align="center">添加产品类别</td>
            </tr>
          </center>
          <tr>
            <td width="39%">
              <p align="right">添加：</td>
            <center>
            <td width="61%"><input type="text" name="pic_class" size="20"></td>
            </tr>
          </center>
          <tr>
            <td width="100%" colspan="2">
              <p align="right"><input type="submit" value="提交" name="B1"><input type="reset" value="全部重写" name="B2"></td>
          </tr>
        </table>
        <center>
        <p>　</p>
        </form>
        <p>　</td>
    </tr>
  </table>
  </center>
</div>


</body>

