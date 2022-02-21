<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
td { font-size: 9pt; line-height: 13pt; text-decoration: none}
-->
</style>
</head>
<body topmargin="0" bgcolor="#e3f1d1">
<table width="100%" border=0 cellspacing=0 cellpadding=0>
<form name="form" method="post" action="upfilesave.asp" enctype="multipart/form-data">
<tr><td  valign=top height=22>
<input type="hidden" name="filepath" value="UploadFile/">
<input type="hidden" name="act" value="upload">
<input type="hidden" name="fname">
<input type="file" name="file1">
<input type="submit" name="Submit" value="上传" onclick="fname.value=file1.value,parent.document.forms[0].Submit.disabled=true,parent.document.forms[0].Submit2.disabled=true;">
限制每个文件2048K以内
</td></tr>
</form></table>
</body>
</html>