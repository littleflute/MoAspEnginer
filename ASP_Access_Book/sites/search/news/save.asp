<%
if request.cookies("adminok")="" then
  response.redirect "login.asp"
end if
%>
<%
if request.form("typename")="" then 
response.write "错误提示:未输入信息类型!"
response.end
end if
if request.form("vote")="" then 
response.write "错误提示:未做评价！"
response.end
end if
%>
<!--#include file="articleconn.asp"-->
<!--#include file="inc/articlechar.inc"-->
<%
dim listname
dim typename
dim title
dim content
dim sql
dim rs
dim filename
dim articleid
dim outfile
dim url
dim from
dim fromurl
dim big
dim vote

title=htmlencode2(request.form("txttitle"))
url=htmlencode2(request.form("txturl"))
content=htmlencode2(request.form("txtcontent"))
typename=htmlencode2(request.form("typename"))
from=htmlencode2(request.form("from"))
fromurl=htmlencode2(request.form("fromurl"))
big=htmlencode2(request.form("big"))
vote=htmlencode2(request.form("vote"))

set rs=server.createobject("adodb.recordset")
sql="select * from learning where (articleid is null)" 

rs.open sql,conn,1,3
rs.addnew
rs("title")=title
rs("url")=url
rs("content")=content
rs("type")=typename
rs("big")=big
rs("from")=from
rs("fromurl")=fromurl
rs("vote")=vote
rs("dateandtime")=date()
rs.update

articleid=rs("articleid")

rs.close
set rs=noting
conn.close
set conn=nothing

%><head>
<link rel="stylesheet" href="css/article.css">
</head>



<div align="center">
  <p>&nbsp;</p><table border="1" cellspacing="0" width="50%" bgcolor="#F0F8FF" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
    <tr>
      <td width="100%" bgcolor="#B5D85E" height="20"> 
        <p align="center"><font color="#FFFFFF"><b>添加信息成功</b></font> 
      </td>
    </tr>
    <tr>
      <td width="100%" height="177"> 
        <table width="80%" border="0" cellspacing="0" cellpadding="2" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
      
          <tr>
            <td>信息名称为：<font color="#FF0000"><%response.write title%></font></td>
          </tr>
        </table>
        <p align="center">&nbsp;</p>
        <p align="center"> <a href="add.asp"> 继续添加数据</a>&nbsp;&nbsp; <a href="manage.asp">返回管理界面</a><br>
        </p>
      </td>
    </tr>
    </table>
</div>
