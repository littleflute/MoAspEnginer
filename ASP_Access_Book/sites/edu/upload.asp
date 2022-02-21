<!--#include file="conn.asp"-->
<!--#include file="isteacher.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE>TD {
	FONT-SIZE: 9pt; LINE-HEIGHT: 140%
}
BODY {
	FONT-SIZE: 9pt; LINE-HEIGHT: 140%
}
A:link {
	COLOR: #0033cc; TEXT-DECORATION: none
}
A:visited {
	COLOR: #0033cc; TEXT-DECORATION: none
}
A:active {
	COLOR: #ff0000; TEXT-DECORATION: none
}
A:hover {
	COLOR: #000000; TEXT-DECORATION: underline
}
</STYLE>
<title>资料上传</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
  <table border="0"  cellspacing="0" cellpadding="0" width=100%>
    <tr>
      <td>
<form method="post" action="upfile.asp" enctype="multipart/form-data" >
<input type="hidden" name="fname">
<input type="file" name="file1" size=35>
<input type="submit" name="Submit" value="上传" onclick="fname.value=file1.value,parent.document.forms[0].submit.disabled=true;">
</form>
</td></tr>
</table>
</body>
</html>
<%
conn.close
set conn = nothing
%>