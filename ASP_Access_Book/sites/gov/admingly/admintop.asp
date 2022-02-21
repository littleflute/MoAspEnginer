<!--#include file="rscoon_1.asp"-->
<html>
<head>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../xsmz.css" rel=stylesheet>
<script language="javascript">
function BoardWin(url) {
  var oth="toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,left=200,top=200";
  oth = oth+",width=500,height=300";
  var BoardWin = window.open(url,"BoardWin",oth);
  BoardWin.focus();
  return false;
}

function SelectChk(url)
{
  var s = false; //用来记录是否存在被选中的复选框
  var Boardid, n=0; 
  var strid, strurl;
  var nn = self.document.all.item("Board"); //返回复选框Board的数量
  for (j=0; j<nn.length; j++) {
    if (self.document.all.item("Board",j).checked) {
      n = n + 1;
      s = true;
      Boardid = self.document.all.item("Board",j).id+"";  //转换为字符串
      //生成要删除公告编号的列表
      if(n==1) {
        strid = Boardid;
      }
      else {
        strid = strid + "," + Boardid;
      }
    }
  }
  strurl = url+"?id=" + strid;
  if(!s) {
    alert("请选择要删除的文件!");
    return false;
  }	
  if (confirm("你确定要删除这些文件吗？")) {
    form1.action = strurl;
	form1.submit();
  }
}

function sltAll()
{
	var nn = self.document.all.item("Board");
	for(j=0;j<nn.length;j++)
	{
		self.document.all.item("Board",j).checked = true;
	}
}
function sltNull()
{
	var nn = self.document.all.item("Board");
	for(j=0;j<nn.length;j++)
	{
		self.document.all.item("Board",j).checked = false;
	}
}
</script>
<script>
oncontextmenu="window.event.returnvalue=false"
</script>
</head>
<body bgcolor="#EFEFEF" background="admin_top_bg.gif" oncontextmenu=return(false)>
<TABLE height="0%" cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
    <TR vAlign=center>
      <TD width=31><img src="admin_top_close.gif" width="31" height="31" style="CURSOR: hand" title=关闭左边管理菜单 
      onClick=switchBar(this)></TD>
      
    <TD width=69><a href="../index.asp" target="_blank"><font color="#FFFFFF">回首页</font></a> 
    </TD>
      <TD width=40><IMG src="admin_top_icon_1.gif" width="30" height="30"> </TD>
      
    <TD width=100><a href="help.asp"><font color="#FFFFFF">【使用帮助】</font></a></TD>
      <TD width=40><IMG src="admin_top_icon_5.gif" width="30" height="30"> </TD>
      
    <TD width=100><a href="#" onclick="BoardWin('mail.asp')"><font color="#FFFFFF">【问题反馈】</font></a> 
    </TD>
      <TD width=40><IMG height=30 
      src="admin_top_icon_6.gif" width=30></TD>
      
    <TD width=112><a href="listalladmin.asp"><font color="#FFFFFF">【管理员名单】</font></a></TD>
      <!--
	<td width=40>
		<img src="images/admin_top_icon_2.gif">
	</td>
	<td width=100>
		<a href="#">我的通讯录</a>
	</td>
	<td width=40>
		<img src="images/admin_top_icon_3.gif">
	</td>
	<td width=100>
		<a href="#">我的备忘录</a>
	</td>
	<td width=40>
		<img src="images/admin_top_icon_4.gif">
	</td>
	<td width=100>
		<a href="#">我的短信息</a>
	</td>
-->
      
    <TD width="471" align=middle> 
      <div align="left"><a href="<%response.write(script_name&"?"&query_string)%>"><font color="#FFFFFF">【刷新】</font></a> 
        <a href="servercheck.asp"><font color="#FFFFFF">【查看服务器属性】</font></a></div>
    </TD>
    </TR>
  </TBODY>
</TABLE>
<TABLE class=border cellSpacing=1 cellPadding=2 width="100%" align=center 
border=0>
  <TBODY>
    <TR align=middle bgcolor="#CCCCCC"> 
      <TD height=25 class=topbg><strong><font color="#FF0000">政府网局网站后台管理系统!</font></strong></TD>
    <TR> 
      <TD height=23 background="topBar_bg.gif" class=tdbg>此后台操作是直接在数据库的操作,所以请你在操作时要小心谨慎,不要因为你一时的误操作而导致不必要的损失!</TD>
    </TR>
  </TBODY>
</TABLE>
<p>&nbsp;</p>
<br>
<br>

