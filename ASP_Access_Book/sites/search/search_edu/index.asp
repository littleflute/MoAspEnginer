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
font-family: Tahoma, Verdana; font-size: 9pt; color: #FFFFFF; background-color: #B05706
}
.category{
font-family: Tahoma, Verdana; font-size: 9pt; color: #000000; background-color: #EFEFEF
}
</STYLE>
<title><%=schoolname%></title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<!--#include file="head.asp"-->
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=5 width="750" border=0>
<tr align=center valign=top>
    <td width=170> <br>
    </td>
    <td width=430> <br>
      <form action="list.asp" method="post">
        <table style="BORDER-COLLAPSE: collapse" bordercolor=#808080 align=center cellpadding=1 width=100% border=1>
          <tr>
            <td align="center" class="header" colspan=2>快速搜索</td>
          </tr>
          <tr> 
            <td width=40 align="center">资料</td>
            <td>&nbsp; 
              <input type=text name="course" size=12>
            </td>
          </tr>
          <tr> 
            <td align="center">标题</td>
            <td>&nbsp; 
              <input type=text name="title" size=12>
            </td>
          </tr>
          <%set rs = server.createobject("adodb.recordset")
sql = "select * from type"
rs.open sql,conn,1,1
%>
          <tr> 
            <td align="center">类型</td>
            <td> &nbsp; 
              <select name="filetype">
                <option value="" selected>请选择&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
                <%
do while not rs.eof
%>
                <option value="<%=rs("typeid")%>"><%=rs("type")%></option>
                <%
rs.movenext
loop
rs.close
%>
              </select>
            </td>
          </tr>
          <tr> 
            <td colspan=2 align=center> 
              <input type=submit name="submit" value="搜索">
              &nbsp;&nbsp; 
              <input type=reset name="reset" value="清空">
            </td>
          </tr>
        </table>
      </form>
    </td>
<td width=150>
    </td>
</tr>
</table>
<!--#include file="foot.asp"-->
<%
set rs = nothing
conn.close
set conn = nothing
%>
</body>
</html>