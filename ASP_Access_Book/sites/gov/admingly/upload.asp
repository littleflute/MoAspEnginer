<html><head>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../xsmz.css" rel=stylesheet></hedad>
<body bgcolor="#EFEFEF">
<form name="frmCtoy" method="post" action="moviesave.asp?action=save" >
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000000">
    <tr> 
      <td height="30" colspan="2" bgcolor="#B4E7F1"> 
        <div align="center"><strong>上传视频文件</strong></div></td></tr>
    <tr>
      <td width="59%">视频文件标题:</td>
      <td height="25"> <input name="vdtitle" type="text" size="30"></td>
    <tr>
	<tr>
      <td width="59%">内容简介:</td>
      <td width="41%"> 
        <textarea name="introvd" cols="30" rows="5"></textarea>
      </td>
    </tr>
 <tr> 
      <td height="40">上传文件：</td>
      <td height="25"> 
        <iframe frameborder="0" width="400" height=24 scrolling="no" src="upload1.asp" ></iframe>
     <input type="hidden" name="filename">
      </td>
  </tr>
  <tr>
      <td height="25">&nbsp;</td>
    <td><input type="submit" name="ups" value="确定"></td>
	</tr>
</table>
</form>
</body>
</html>