<!--#include file="admintop.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%  dim filename,i
  id = Request.QueryString("id")
  sql="select movies from video where id in("&id&")"
    set rs=conn.execute(sql)
	      filename=rs.getrows
  sql = "DELETE FROM video WHERE id In (" & id & ")"
  Conn.Execute(sql)
   for i=0 to ubound(filename,2)
   response.write"<script>alert('"&delbackdb(filename(0,i))&"')</script>"
   next
%>
<script language="javascript">
  alert("数据库文件已成功删除！");
</script>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td height="30" bgcolor="#B1E1F3"> 
      <div align="center"><strong>删除成功</strong></div></td>
  </tr>
  <tr>
    <td><div align="center">删除的文件将是永久删除,操作时请一定谨慎对待.请选择你返回要返回的页面.</div></td>
  </tr>
  <tr>
    <td height="20">
<div align="center"><a href="video.asp">【返回】</a><a href="adminbody.asp">【继续其它操作】</a></div></td>
  </tr>
</table>

</body>
</html>
