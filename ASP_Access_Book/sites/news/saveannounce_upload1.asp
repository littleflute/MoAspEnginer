<script>
if (top.location==self.location){
	top.location="index.asp"
}
</script>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link rel="stylesheet" type="text/css" href="style.css">
<body  class="tdbg" text="#000000" leftmargin="0" topmargin="0">
<form name="form" method="post" action="saveannouce_upfile1.asp" enctype="multipart/form-data" >
<!-- 注意文件上传路径 -->
<input type="hidden" name="filepath" value="upfile/">
<input type="hidden" name="act" value="upload">
<input type="file" name="file1" size=20 class="smallinput">
<input type="submit" name="Submit" value="上传" class="but"> 类型：*.gif、*.jpg、*.doc，限制：50K
</form>
</body>
</html>