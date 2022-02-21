<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css">
<script language="javascript">
function showinfo()
{
lay1.style.visibility="visible";
}
</script>
</head>
<div id="lay1" style="visibility:hidden;position:absolute;width:260px;height:90px;z-index:2;background:#ffffff;padding:12px;line-height:12px">上传中....</div>
<body  leftmargin="0" topmargin="0" >
<form name="form1" method="post" action="upfile.asp" enctype="multipart/form-data" >
<input type="hidden" name="act" value="upload">
<input type="hidden" name="filepath" value="../images">
  <table width="100%" height="100%" border="0" align="center" cellspacing="0" bordercolorlight="#000000">
    <tr>
<td>
<input type="file" name="file1" style="width:88%"  class="tx1" value="">
</td><td><input type="submit" name="Submit" value="上传" onClick="showinfo()">
</td>
</tr>
</table>
</form>
</body>
</html>
