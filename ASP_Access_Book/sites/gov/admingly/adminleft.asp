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
      <td background="Admin_left_1.gif"><font color="#000000"><a href="admin_sort.asp" target="main">��Ŀ�������</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="sort_new.asp" target="main">��Ŀ���Ź���</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_policy.asp" target="main">���������ļ�</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="video.asp" target="main">��Ƶ�ļ�����</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_vote.asp" target="main">ͶƱͳ�ƹ���</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_tj.asp" target="main">ͳ�Ʊ����</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_news.asp" target="main">�������Ź���</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_worknet.asp" target="main">���ϰ칫�ļ�����</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_consul.asp" target="main">�û���ѯ����</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_link.asp" target="main">�������ӹ���</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_basic.asp" target="main">������������</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_top.asp" target="main">������Ϣ����</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="admin_affiche.asp" target="main">�����趨</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="addadmin.asp" target="main">��ӹ���Ա</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="editpwd.asp" target="main">��������</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="unlockadmin.asp" target="main">����˺�����</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="unlock.asp" target="main">���IP����</a></font></td>
    </tr>
    <tr> 
      <td background="Admin_left_1.gif"><font color="#FFFFFF"><a href="loginout.asp" target="main">�˳�����</a></font></td>
    </tr>
  </table>
</div>
</body>
</html>
