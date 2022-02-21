<html>
<head>
<LINK href="style.css" rel=stylesheet>
</head>
<body>
<%
if session("admin")="" and session("flag")="" then
	response.write "ÇëµÇÂ½ÏÈ"
else
	call main()
end if
sub main()
%><br><br>
<table border=0 width=100% cellspacing="2" cellpadding="1">
<tr>
    <td width=100% align=center> 
      <p><font style="font-size: 14.8px;line-height:150%"> <img src="../a.jpg" width="154" height="295" usemap="#Map" border="0"></font></p>
      <p><font style="font-size: 14.8px;line-height:150%"><br>
        <br>
        </font> </p></td>
  </tr></table>
<%
end sub
%><body>
<map name="Map">
  <area shape="rect" coords="38,51,126,73" href="adminedit.asp" target="list">
  <area shape="rect" coords="34,102,128,130" href="classmana.asp" target="list">
  <area shape="rect" coords="36,151,128,180" href="adminuser.asp" target="list">
  <area shape="rect" coords="37,207,127,231" href="main.asp" target="list">
  <area shape="rect" coords="38,257,130,284" href="logout.asp" target="list">
</map>
</html>