<!--#include file="rscoon_1.asp"-->
<!--#include file="isadmin.asp"-->
<html>
<head>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../xsmz.css" rel=stylesheet>
</head>
<body>
<br><br>
<form name="form1" method="post" action="addvote.asp">
  <table width="400" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000" bgcolor="#EFEFEF">
    <tr bgcolor="#B1E1F3"> 
      <td colspan="2"> 
        <div align="center"><strong>添加投票标题</strong></div></td>
  </tr>
  <tr> 
      <td>投票标题:</td>
    <td><input type="text" name="title"></td>
  </tr>
  <tr> 
      <td>投票项目个数:</td>
    <td><input name="votecount" type="text" size="4"></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><input type="submit" name="Submt" value="下一步"></td>
  </tr>
</table>
</form>
</body>
</html>
