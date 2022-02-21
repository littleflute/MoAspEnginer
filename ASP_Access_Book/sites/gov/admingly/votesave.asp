<!--#include file="rscoon_1.asp"-->
<!--#include file="isadmin.asp"-->
<html>
<head>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../xsmz.css" rel=stylesheet>
</head>
<body>
<%
     dim allcount,id,i,errorcount
	   errorcount=0
	 allcount=trim(request.querystring("allcount"))
	 id=trim(request.querystring("id"))
         dim votes(9)
	 if allcount="" or id="" then
	 response.write"<script>alert('对不起,出现严重错误');history.back();</script>"
           errorcount=1
		   end if
     for i=1 to allcount
	 if trim(request.form("vote"&i))="" then
    response.write"<script>alert('请把投票项目填写完整');history.back();</script>"
      errorcount=1
         else
		 votes(i)=trim(request.form("vote"&i))
		  end if
		  next
		  if errorcount=0 then
		  for i=1 to allcount
		  sql="insert into votes(votetitle,fromid) values('"&votes(i)&"',"&id&")"
		     conn.execute(sql)
			 next
%>
<table width="100%" border="0" align="center" bgcolor="#EFEFEF">
  <tr>
    <td bgcolor="#B1E1F3">
<div align="center"><strong>添加成功</strong></div></td>
  </tr>
  <tr>
    <td>请关闭窗口,在父窗口点击【刷新】再点击您所添加的项,查看详情!</td>
  </tr>
  <tr>
    <td><div align="center"><a href="javascript:window.close()">【关闭窗口】</a></div></td>
  </tr>
</table>
<% end if %>
</body>
</html>
