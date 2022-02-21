<!--#include file="rscoon_1.asp"-->
<!--#include file="isadmin.asp"-->
<html>
<head>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../xsmz.css" rel=stylesheet>
</head>
<% dim id,errorcount
     errorcount=""
id=request.querystring("id")
  sql="select isvote from stat where id="&id
Set rs = Server.CreateObject("ADODB.RecordSet")
         rs.open sql,conn,1,3
		 if rs("isvote")="true" then
		 rs("isvote")="false"
		 else
		 sql1="select isvote,title from stat where isvote='true'"
		 set rs1=conn.execute(sql1)
		 if not rs1.eof then
		 response.write"<script>alert('对不起,已经开启了别的投票了\n要想开启此投票统计,请先关闭已开启的那个投票选项');</script>"
		    errorcount="已经有【"&rs1("title")&"】投票统计开启了,请先关闭此选项,才能打开其它的投票选项"
		    else
		 rs("isvote")="true"
		 end if 
		 end if
		 rs.update
 if errorcount="" then %>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="30" bgcolor="#B1E1F3"> <div align="center"><strong><% if rs("isvote")="true" then response.write"开启" else response.write"关闭" end if%>成功</strong></div></td>
  </tr>
  <tr> 
    <td><div align="center">请关闭窗口,然后点击父窗口的【刷新】进行查看.</div></td>
  </tr>
  <tr> 
    <td height="20"> <div align="center"><a href="javascript:window.close()">【关闭窗口】</a></div></td>
  </tr>
</table>
<% else %>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="30" bgcolor="#B1E1F3"> <div align="center"><strong><% if rs("isvote")="true" then response.write"关闭" else response.write"开启" end if%>失败</strong></div></td>
  </tr>
  <tr> 
    <td><div align="center"><%=errorcount%></div></td>
  </tr>
  <tr> 
    <td height="20"> <div align="center"><a href="javascript:window.close()">【关闭窗口】</a></div></td>
  </tr>
</table>
<% end if %>
</body>
</html>
