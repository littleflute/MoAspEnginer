<html>
<head>
<title>添加文件</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../style.css" type="text/css">
</head>

<body background="bkg.gif">
<div align="center">注意：本功能是基于WEB方式的上传文件，比使用FTP方便!文件上传在站点的根目录<font color="#FF0000">img/</font>目录下。<br>
  支持上传的文件格式有：图片文件（*.gif、*.jpg、*.png、*.bmp），动画FLASH文件（*.swf）<br>
  调用图片请用绝对路径也就是如例：<font color="#FF0000">http://你的网址/img/1.gif</font>，这样就调用了<font color="#FF0000">1.gif</font>这张图片<br>
  当您填加的一个栏目需要图片的时候,请先上传该图片,并复制该图片地址,粘贴到相应栏目的需求图片位置上即可<br>
  <span class="font"><br>
  <font color="#FF0000">注意：请不要用此功能上传大于2M的文件。如果你要上传的文件大于2M，请用FTP上传！<br>
  1M等于1024K，1K等于1024个字节<br>
  </font></span></div>
<form method="post" action="sub_upload.asp" enctype="multipart/form-data">
  <table border=0 cellspacing=0 cellpadding=0 align=center width="333">
    <tr> 
<td height=20 width=60>&nbsp;文件</TD>
<td height="16" width="371"> 
<input type="file" name="src" size="20" value="浏览">
&nbsp;&nbsp; 
<input type="submit" value="上传" name="B1" IsShowProcessBar="True">
</td>
</tr>
</table>
</form>
</body>
</html>