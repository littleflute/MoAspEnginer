<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>政府网后台管理系统</title>
<script LANGUAGE="JavaScript">
<!--
// Scrolling message settings
var MessageText = "民政后台管理"
var DisplayLength = 260
var pos = 1 - DisplayLength;
function ScrollInStatusBar(){
var scroll = "";
pos++;
if (pos == MessageText.length) pos = 1 - DisplayLength;
if (pos<0) {
for (var i=1; i<=Math.abs(pos); i++)
scroll = scroll + "";
scroll = scroll + MessageText.substring(0, DisplayLength - i + 1);
}
else
scroll = scroll + MessageText.substring(pos, pos + DisplayLength);
window.status = scroll;
//Scrolling speed
setTimeout("ScrollInStatusBar()",50);
}
ScrollInStatusBar()
//-->
</script>
</head>
<frameset rows="*" cols="180,*" framespacing="0" frameborder="no" border="0">
  <frame src="adminleft.asp" name="leftFrame" scrolling="yes" noresize>
  <frame src="adminbody.asp" name="main">
</frameset>
<noframes><body>

</body></noframes>
</html>
