<% Time1=Timer %>
<!--#include file="odbc_connection.asp"-->
<!--#include file="function.asp"-->
<html>
	<head>
		<title>小论坛</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="image/bbs.css" rel=stylesheet>
<style type="text/css">
<!--
a:visited{text-transform: none; text-decoration: none; font-style: normal; font-weight: normal; font-variant: normal; color: #000000}
a:hover{text-decoration:underline; color: #FF3333}
a:link{text-transform: none; text-decoration: none; font-weight: normal; font-variant: normal; color: #000000}
.p9 { font-size: 9pt; line-height: 13pt; text-decoration: none}
.p12 { font-size: 14px; line-height: 13pt; text-decoration: none}

td { font-size: 9pt; line-height: 13pt; text-decoration: none}
.pic_show { filter:dropshadow(color=#999999,offx=5,offy=5.positive=1)}
.ptitle_9 { font-size: 9pt; line-height: 11pt}
.pbutton { border: #666666; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
body {
	margin-top: 0px;
}
-->
</style>
	</head>
<body background="pic/lvbgcolor.gif">
<!-- #include file="menu.asp" -->
<TABLE width="760" align=center border=0>
  <TBODY>
    <TR> 
      <TD><img src="a.jpg" width="759" height="116"></TD>
    </TR>
  </TBODY>
</TABLE>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr align="center"> 
    <td width="760" bgcolor="#CCCCCC"> 
      <%  
sqlgg="select * from news where boardid= 0 and layer=1 order by submit_date desc"
set rsgg = db.execute(sqlgg)
%>
      <SCRIPT LANGUAGE='JavaScript' SRC='js/fader.js' TYPE='text/javascript'></script> 
      <SCRIPT LANGUAGE='JavaScript' TYPE='text/javascript'>
prefix="";
arNews = ["<% If not (rsgg.eof and rsgg.bof)Then %><a href='particular.asp?bbs_id=<%= rsgg("bbs_id") %>&boardid=0' target=_blank><b><%= rsgg("title") %></b></a> [<%= rsgg("submit_date") %>]<% Else %>当前还没有公告<% End If %>",""," <% If not rsgg.eof Then
 rsgg.movenext
If not rsgg.eof Then %><a href='particular.asp?bbs_id=<%= rsgg("bbs_id") %>&boardid=0' target=_blank><b><%= rsgg("title") %></b></a> [<%= rsgg("submit_date") %>]<% Else %>论坛欢迎你的光临<% End If %> <% Else %>论坛欢迎你的光临<% End If %>","",
" <% If not rsgg.eof Then
 rsgg.movenext
If not rsgg.eof Then %><a href='particular.asp?bbs_id=<%= rsgg("bbs_id") %>&boardid=0' target=_blank><b><%= rsgg("title") %></b></a> [<%= rsgg("submit_date") %>]<% Else %>严禁恶意使用粗言秽语，违者经劝告无效，立即封ID<% End If %>
<% Else %>严禁恶意使用粗言秽语，违者经劝告无效，立即封ID<% End If %>",""]
</SCRIPT> <div id="elFader" style="position:relative;visibility:hidden; height:16" ></div> 
      <% 
rsgg.close
set rsgg=nothing
 %>
  </tr>
  <tr > 
    <td height=1 bgcolor="#a9d46d"> </tr>
  <tr> 
    <td height="20" bgcolor="#CCCCCC"> <!--#include file="menu_include.asp"--> </td>
  </tr>
</table>
<!--#include file="login_include.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid;">
  <tr>
    <td height="23" colspan="3" background="pic/h3.gif" class="p12"><img src="c.jpg" width="759" height="25"></td>
  </tr>
  <tr> 
    <td height="23" colspan="3" background="pic/h3.gif" class="p12"><strong> <font color="#FFFFFF">-=■=-专区</font></strong></td>
  </tr>
  <tr> 
    <td width="80" rowspan="2" align="center" valign="middle" bgcolor="#E1FAFF"><img src="z11.gif" width="45" height="45"></td>
    <td width="469" height="36" bgcolor="#F4FDFF" style="border-left: 1 solid #a9d46d">&nbsp;&nbsp;<span class="p12"><a href="board.asp?boardid=1"><font color="#FF0000">专区 
      </font></a></span><br>
      <font color="#000000"><img src="pic/Forum_readme.gif" width="10" height="10">关于本站的一切问题，对本站的意见建议等</font></td>
    <td width="209" height="36" bgcolor="#F4FDFF"> 
      <%
   set board1rs=Server.CreateObject("ADODB.Recordset")
   board1sql="select * from news where boardid=1 order by submit_date desc"
   board1rs.open board1sql,db,1,3
 %>
      主题：<%= gotTopic(board1rs("title"),20) %><br>
      作者：<%= board1rs("user_name") %><br>
      日期：<%= board1rs("submit_date") %> 
      <% 
board1rs.close
set board1rs=nothing
%>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#E1FAFF" style="border-left: 1 solid #a9d46d">版主： 
      <% 
	if session("banzhu1")="" then
	set rsbanzhu1 = db.execute("select banzhu from board where boardid=1")
	session("banzhu1")=rsbanzhu1("banzhu")
	rsbanzhu1.close
	set rsbanzhu1=nothing
	end if
	Response.Write(session("banzhu1"))
	%>
    </td>
    <td bgcolor="#E1FAFF"> <font color="#666666">今日</font>：<font color="#FF0000">0</font> 
      ,<font color="#666666">主题</font>： , <font color="#666666">帖子</font>：</td>
  </tr>
</table>
<img src="" alt="" name="" width="32" height="4" style="background-color: #FFFFFF"> 
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid;">
  <tr>
    <td height="23" colspan="3" background="pic/h3.gif" class="p12"><img src="c.jpg" width="759" height="25"></td>
  </tr>
  <tr> 
    <td height="23" colspan="3" background="pic/h3.gif" class="p12"><strong> <font color="#FFFFFF">-=■=-创业论坛</font></strong></td>
  </tr>
  <tr bgcolor="#E1FAFF"> 
    <td width="80" rowspan="2" align="center" valign="middle"><img src="z6.gif" width="45" height="45"></td>
    <td width="469" height="36" bgcolor="#F4FDFF" style="border-left: 1 solid #a9d46d">&nbsp;&nbsp;<span class="p12"><a href="board.asp?boardid=2"><font color="#FF0000">创业论坛</font></a></span><br>
      <font color="#000000"><img src="pic/Forum_readme.gif" width="10" height="10">关于本站的一切问题，对本站的意见建议等</font></td>
    <td width="209" height="36" bgcolor="#F4FDFF"> 
      <%
   set board1rs=Server.CreateObject("ADODB.Recordset")
   board1sql="select * from news where boardid=2 order by submit_date desc"
   board1rs.open board1sql,db,1,3
 %>
      主题：<%= gotTopic(board1rs("title"),20) %><br>
      作者: <%= board1rs("user_name") %> <br>
      日期: <%= board1rs("submit_date") %> 
      <% 
board1rs.close
set board1rs=nothing
%>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#E1FAFF" style="border-left: 1 solid #a9d46d">版主： 
      <% 
	if session("banzhu2")="" then
	set rsbanzhu2 = db.execute("select banzhu from board where boardid=2")
	session("banzhu2")=rsbanzhu2("banzhu")
	rsbanzhu2.close
	set rsbanzhu2=nothing
	end if
	Response.Write(session("banzhu2"))
	%>
    </td>
    <td bgcolor="#E1FAFF"><font color="#666666">今日</font>：<font color="#FF0000">0</font> 
      ,<font color="#666666">主题</font>： , <font color="#666666">帖子</font>：</td>
  </tr>
</table>
<img src="" alt="" name="" width="32" height="4" style="background-color: #FFFFFF"> 
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid;border-bottom: 1 solid #a9d46d">
  <tr>
    <td height="23" colspan="3" background="pic/h3.gif"><img src="c.jpg" width="759" height="25"></td>
  </tr>
  <tr> 
    <td height="23" colspan="3" background="pic/h3.gif"><strong> <span class="p12"><font color="#FFFFFF">-=■=-娱乐休闲</font></span></strong></td>
  </tr>
  <tr> 
    <td width="80" rowspan="2" align="center" valign="middle" bgcolor="#E1FAFF"  style="border-bottom: 1 solid #a9d46d"><img src="z9.gif" width="45" height="45"></td>
    <td width="469" height="36" bgcolor="#F4FDFF" style="border-left: 1 solid #a9d46d">&nbsp;&nbsp;<span class="p12"><a href="board.asp?boardid=3"><font color="#FF0000">贴图专区<br>
      </font></a></span> <font color="#000000"><img src="pic/Forum_readme.gif" width="10" height="10">关于本站的一切问题，对本站的意见建议等</font></td>
    <td width="209" height="36" bgcolor="#F4FDFF"> 
      <%
   set board1rs=Server.CreateObject("ADODB.Recordset")
   board1sql="select * from news where boardid=3 order by submit_date desc"
   board1rs.open board1sql,db,1,3
 %>
      主题：<%= gotTopic(board1rs("title"),20) %><br>
      作者: <%= board1rs("user_name") %> <br>
      日期: <%= board1rs("submit_date") %> 
      <% 
board1rs.close
set board1rs=nothing
%>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#E1FAFF" style="border-left: 1 solid #a9d46d;border-bottom: 1 solid #a9d46d">版主： 
      <% 
	if session("banzhu3")="" then
	set rsbanzhu3 = db.execute("select banzhu from board where boardid=3")
	session("banzhu3")=rsbanzhu3("banzhu")
	rsbanzhu3.close
	set rsbanzhu3=nothing
	end if
	Response.Write(session("banzhu3"))
	%>
    </td>
    <td bgcolor="#E1FAFF" style="border-bottom: 1 solid #a9d46d"><font color="#666666">今日</font>：<font color="#FF0000">0</font> 
      ,<font color="#666666">主题</font>： , <font color="#666666">帖子</font>：</td>
  </tr>
  <tr> 
    <td width="80" rowspan="2" align="center" bgcolor="#E1FAFF"><img src="z16.gif" width="45" height="45"></td>
    <td width="469" height="36" bgcolor="#F4FDFF" style="border-left: 1 solid #a9d46d">&nbsp;&nbsp;<span class="p12"><a href="board.asp?boardid=4"><font color="#FF0000">灌水天堂<br>
      </font></a></span> <font color="#000000"><img src="pic/Forum_readme.gif" width="10" height="10">关于本站的一切问题，对本站的意见建议等</font></td>
    <td width="209" height="36" bgcolor="#F4FDFF"> 
      <%
   set board1rs=Server.CreateObject("ADODB.Recordset")
   board1sql="select * from news where boardid=4 order by submit_date desc"
   board1rs.open board1sql,db,1,3
 %>
      主题：<%= gotTopic(board1rs("title"),20) %><br>
      作者: <%= board1rs("user_name") %> <br>
      日期: <%= board1rs("submit_date") %> 
      <% 
board1rs.close
set board1rs=nothing
%>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#E1FAFF" style="border-left: 1 solid #a9d46d">版主： 
      <% 
	if session("banzhu4")="" then
	set rsbanzhu4 = db.execute("select banzhu from board where boardid=4")
	session("banzhu4")=rsbanzhu4("banzhu")
	rsbanzhu4.close
	set rsbanzhu4=nothing
	end if
	Response.Write(session("banzhu4"))
	%>
    </td>
    <td bgcolor="#E1FAFF"><font color="#666666">今日</font>：<font color="#FF0000">0</font> 
      ,<font color="#666666">主题</font>： , <font color="#666666">帖子</font>：</td>
  </tr>
</table>
<% Time2=Timer %>
<TABLE class=a2 cellSpacing=1 cellPadding=0 width="760" align=center 
  border=0>
  <TBODY>
    <TR>
      <TD height=25 colSpan=2 background="pic/h3.gif" ><img src="c.jpg" width="759" height="25"></TD>
    </TR>
    <TR> 
      <TD height=25 colSpan=2 background="pic/h3.gif" >　<B>■ </B>在线统计</TD>
    </TR>
    <TR> 
      <TD class=a3 align=middle width="5%"><IMG 
      src="image/totel.gif"></TD>
      <TD class=a4 vAlign=top><TABLE cellSpacing=0 cellPadding=3 width="100%">
          <TBODY>
            <TR> 
              <TD height=15>&nbsp;<IMG id=followImg0 style="CURSOR: hand" 
            onclick=loadThreadFollow(0,0) 
            src="image/plus.gif" loaded="no"> 目前论坛总共 有 <B>1</B> 人在线。其中注册用户 <B>0</B> 
                人，访客 <B>1</B> 人。最高在线 <FONT 
            color=red><B>4</B></FONT> 人，发生在 <B>2006-5-5 19:59:10</B> </TD>
            </TR>
            <TR id=follow0 style="DISPLAY: none" height=25> 
              <TD class=a4 id=followTd0 align=left width="94%" 
            colSpan=5>　Loading.....</TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
  </TBODY>
</TABLE>
<br>
<TABLE class=a2 cellSpacing=1 cellPadding=1 width="760" align=center 
  border=0>
  <TBODY>
    <TR>
      <TD height=25 colSpan=2 background="pic/h3.gif" ><img src="c.jpg" width="759" height="25"></TD>
    </TR>
    <TR> 
      <TD height=25 colSpan=2 background="pic/h3.gif" >　■<B> </B>友情链接<span class="a4"><a 
      title="天津宠物网&#10;天津宠物门户网站" href="http://www.022pet.com/" 
      target=_blank></a></span></TD>
    </TR>
    <TR> 
      <TD class=a3 align=middle width="5%" rowSpan=2><IMG 
      src="image/shareforum.gif"></TD>
      <TD class=a4></TD>
    </TR>
    <TR> 
      
    <TD class=a4>&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
<!--#include file="foot.asp"-->
<script language="javascript">
function openUser(id) {
	var Win = window.open("dispuser.asp?name="+id,"openScript","width=350,height=300,resizable=1,scrollbars=1,menubar=0,status=1" );
}

function openScript(url, width, height){
	var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}

function openDis(bid,rid,id){
	self.location="dispbbs.asp?boardid="+bid+"&RootID="+rid+"&id="+id
}

function PopWindow()
{ 
openScript('messanger.asp?action=newmsg',420,320); 
}
var nn = !!document.layers;
var ie = !!document.all;

if (nn) {
  netscape.security.PrivilegeManager.enablePrivilege("UniversalSystemClipboardAccess");
  var fr=new java.awt.Frame();
  var Zwischenablage = fr.getToolkit().getSystemClipboard();
}

function submitonce(theform){
//if IE 4+ or NS 6+
if (document.all||document.getElementById){
//screen thru every element in the form, and hunt down "submit" and "reset"
for (i=0;i<theform.length;i++){
var tempobj=theform.elements[i]
if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
//disable em
tempobj.disabled=true
}
}
}
</script>