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
  var s = false; //������¼�Ƿ���ڱ�ѡ�еĸ�ѡ��
  var Boardid, n=0; 
  var strid, strurl;
  var nn = self.document.all.item("Board"); //���ظ�ѡ��Board������
  for (j=0; j<nn.length; j++) {
    if (self.document.all.item("Board",j).checked) {
      n = n + 1;
      s = true;
      Boardid = self.document.all.item("Board",j).id+"";  //ת��Ϊ�ַ���
      //����Ҫɾ�������ŵ��б�
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
    alert("��ѡ��Ҫɾ�����ļ�!");
    return false;
  }	
  if (confirm("��ȷ��Ҫɾ����Щ�ļ���")) {
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
      <TD width=31><img src="admin_top_close.gif" width="31" height="31" style="CURSOR: hand" title=�ر���߹���˵� 
      onClick=switchBar(this)></TD>
      
    <TD width=69><a href="../index.asp" target="_blank"><font color="#FFFFFF">����ҳ</font></a> 
    </TD>
      <TD width=40><IMG src="admin_top_icon_1.gif" width="30" height="30"> </TD>
      
    <TD width=100><a href="help.asp"><font color="#FFFFFF">��ʹ�ð�����</font></a></TD>
      <TD width=40><IMG src="admin_top_icon_5.gif" width="30" height="30"> </TD>
      
    <TD width=100><a href="#" onclick="BoardWin('mail.asp')"><font color="#FFFFFF">�����ⷴ����</font></a> 
    </TD>
      <TD width=40><IMG height=30 
      src="admin_top_icon_6.gif" width=30></TD>
      
    <TD width=112><a href="listalladmin.asp"><font color="#FFFFFF">������Ա������</font></a></TD>
      <!--
	<td width=40>
		<img src="images/admin_top_icon_2.gif">
	</td>
	<td width=100>
		<a href="#">�ҵ�ͨѶ¼</a>
	</td>
	<td width=40>
		<img src="images/admin_top_icon_3.gif">
	</td>
	<td width=100>
		<a href="#">�ҵı���¼</a>
	</td>
	<td width=40>
		<img src="images/admin_top_icon_4.gif">
	</td>
	<td width=100>
		<a href="#">�ҵĶ���Ϣ</a>
	</td>
-->
      
    <TD width="471" align=middle> 
      <div align="left"><a href="<%response.write(script_name&"?"&query_string)%>"><font color="#FFFFFF">��ˢ�¡�</font></a> 
        <a href="servercheck.asp"><font color="#FFFFFF">���鿴���������ԡ�</font></a></div>
    </TD>
    </TR>
  </TBODY>
</TABLE>
<TABLE class=border cellSpacing=1 cellPadding=2 width="100%" align=center 
border=0>
  <TBODY>
    <TR align=middle bgcolor="#CCCCCC"> 
      <TD height=25 class=topbg><strong><font color="#FF0000">����������վ��̨����ϵͳ!</font></strong></TD>
    <TR> 
      <TD height=23 background="topBar_bg.gif" class=tdbg>�˺�̨������ֱ�������ݿ�Ĳ���,���������ڲ���ʱҪС�Ľ���,��Ҫ��Ϊ��һʱ������������²���Ҫ����ʧ!</TD>
    </TR>
  </TBODY>
</TABLE>
<p>&nbsp;</p>
<br>
<br>

