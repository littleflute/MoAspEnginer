<html><head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title></title>
<style type=text/css>
.o1{position:absolute; font-family:Webdings; background-color:#255ca9; font-size:12px; color:#ffffff; cursor:hand;}
</style>
<script language=javascript>
function changeWin(){
	if(parent.forum.cols != "12,*"){
		parent.forum.cols = "12,*";
		document.all.menuSwitch.innerHTML = "<font class=o1>4</font>";
	}
	else{
		parent.forum.cols = "130,*";
		document.all.menuSwitch.innerHTML = "<font class=o1>3</font>";
	}
}
</script>
</head>
<base target="list">
<body  topmargin=0 marginheight=0 leftmargin=0 marginwidth=0 bgcolor=#d0d0c8>
<table width=100% height=100% border=0 cellpadding=0 cellspacing=0>
<tr><td width=100%><iframe id=forumList name=forumList style="height:100%; width:100%; visibility:inherit;" scrolling=no frameborder=0 src="leftmain.asp"></iframe></td>
<td bgcolor=#000000><img src=images/c.gif width=1 height=1><td>
<td bgcolor=#d0d0c8>
	<table width=100% height=100% border=0 cellpadding=0 cellspacing=0>
	<tr><td height=22 onclick=changeWin()><img src=images/c.gif width=10 height=1></td></tr>
	<tr><td onclick=changeWin() height=100% id=menuSwitch><font class=o1>3</font></td></tr>
	</table>
</td><td bgcolor=#000000><img src=images/c.gif width=1 height=1><td>
</tr>
</table>
</body></html>