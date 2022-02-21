<%
  servername=Request.ServerVariables("SERVER_NAME")
  referer=Request.ServerVariables("http_referer")
  if instr(refere,servername)<0 then
  response.redirect"error.htm"
  end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK 
href="../xsmz.css" rel=stylesheet>
</head>
<script>
oncontextmenu="window.event.returnvalue=false"
</script>
<body bgcolor="#B6D7E7" background="admin_top_bg.gif" oncontextmenu=return(false)>
<div align="center">
  <table width="110" border="0" align="left" cellpadding="3" cellspacing="5">
    <tr>
      <td><img src="../images/xsmz.gif" width="118" height="48"></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#000000"><a href="admin_sort.asp" target="main">栏目分类管理</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="sort_new.asp" target="main">栏目新闻管理</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_policy.asp" target="main">国家政策文件</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="video.asp" target="main">视频文件管理</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_vote.asp" target="main">投票统计管理</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_tj.asp" target="main">统计表管理</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_news.asp" target="main">民政新闻管理</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_worknet.asp" target="main">网上办公文件管理</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_consul.asp" target="main">用户咨询管理</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_link.asp" target="main">友情链接管理</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_basic.asp" target="main">基本参数设置</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_top.asp" target="main">基本信息设置</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_affiche.asp" target="main">公告设定</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="addadmin.asp" target="main">添加管理员</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="editpwd.asp" target="main">更改密码</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="unlockadmin.asp" target="main">解除账号锁定</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="unlock.asp" target="main">解除IP锁定</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="loginout.asp" target="main">退出管理</a></font></td>
    </tr>
  </table>
</div>
</body>
</html>
