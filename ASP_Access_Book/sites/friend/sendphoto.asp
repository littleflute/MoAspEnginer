<%Option Explicit%><!--#include file="conn.asp"--><%
dim rs_lar,rs
dim sql
dim i

'�Ѷϴ��û��Ƿ��Ѿ�ע��
Set rs_lar = Server.CreateObject("ADODB.Recordset")
sql="select * from larchives where user_id="& session("user_id")
rs_lar.open sql,conn,3,2

if rs_lar.eof and rs_lar.bof then
    if session("user_id")=1 then response.redirect "notreg.htm"
	 response.redirect "notregist.htm"
	 response.end
end if

%><!--#include file="connpic.asp"--><%

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from pic where user_id=" & session("user_id")
rs.open sql,conn,3,2
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Namo WebEditor v4.0(Trial)">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>��Ƭ�ϴ�</title>
<style>
<MM:BeginLock translatorClass="MM_SSI" type="ssi_include" depFiles="E:\web\homepage\jiaoyou\css.htm" orig="%3C!--#include file=%22css.htm%22--%3E" ><!--
.x:link      { color: white; text-decoration: none }
.x:visited   { color: white; text-decoration: none }
.x:active    { color: red; text-decoration: none }
.x:hover     { color: red;text-decoration:nome}
tr           { font-size: 10pt }
body         { font-size: 10pt }
a:link       { color: blue; text-decoration: none }
a:visited    { color: blue; text-decoration: none }
a:active     { color: red; text-decoration: none }
a:hover      { color: red; text-decoration: none }
--><MM:EndLock>
</style>
<Script language="javascript">
function mysubmit(theform)
{
    if(theform.big.value=="")
    {
    alert("���������ť��ѡ����Ҫ�ϴ���jpg��gif�ļ�!")
    theform.big.focus;
    return (false);
    }
    else
    {
    str= theform.big.value;
    strs=str.toLowerCase();
    lens=strs.length;
    extname=strs.substring(lens-4,lens);
    if(extname!=".jpg" && extname!=".gif")
    {
    alert("��ѡ��jpg��gif�ļ�!");
    return (false);
    }
    }
    return (true);
}
</script>
<title>����ʶ��</title>
<meta name="generator" content="Namo WebEditor v4.0(Trial)">
<style type="text/css">A:visited {
 COLOR: #000000; TEXT-DECORATION: none
}
A:link {
 COLOR: #000000; TEXT-DECORATION: none
}
A:hover {
 COLOR: #0080c0; TEXT-DECORATION: none
}
TD {
 FONT-SIZE: 9pt; COLOR: #000000; FONT-FAMILY: "����"
}
.f1 {
 LINE-HEIGHT: 18px
}
.en {
 FONT-WEIGHT: bold; FONT-FAMILY: "Arial","Verdana"
}
.new {
 FONT-WEIGHT: bold; COLOR: #ff3300; FONT-FAMILY: "Arial"
}
.line {
 LINE-HEIGHT: 19px
}
</style>
<style type="text/css">
<!--
body {
	background-image: url(bihaibg.jpg);
}
-->
</style>
<SCRIPT language=JavaScript>
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  if (theURL != "fuckyou")
 {   window.open(theURL,winName,features);}
}
//-->
</SCRIPT></head>

<body leftmargin="5" topmargin="0">
<table width="772" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="images/top.gif" width="772" height="112" usemap="#Map" border="0"></td>
  </tr>
</table>
<table width="776" height="182" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="47%" height="65" valign="top">

      <form enctype="multipart/form-data" action="addpic.asp" method=post onsubmit="return mysubmit(this)">
        <table border="0" width="100%" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="277" bgcolor="#336699"><font color="#FFFFFF"><b>��Ƭ�ϴ�</b></font></td>
          </tr>
          <tr> 
            <td width="331"> 
              <p align="center">��Ƭ&nbsp; 
                <input type="file" name="big" size="20">
              </p>
            </td>
          </tr>
          <tr> 
            <td width="337"> 
              <p align="center">
                <input type="submit" value="  �ϴ�  " name="B3">
              ��</p>
            </td>
          </tr>
        </table>
      </form>
      <form>
        <table border="0" width="100%" height="13" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="100%" height="10" bgcolor="#336699" colspan="2"><font color="#FFFFFF"><b>�����ϴ�<%=rs.recordcount%>����Ƭ</b></font></td>
          </tr>
          <%for i=1 to rs.recordcount%> 
          <tr> 
            <td width="55%" height="5">��<%=i%>��</td>
            <td width="45%" height="5"> 
              <p align="right"><a href="displaypic.asp?id=<%=rs("id")%>">[�鿴]</a><a href="delpic.asp?id=<%=rs("id")%>">[ɾ��]</a> 
            </td>
          </tr>
          <%rs.movenext%> <%next%> 
        </table>
      </form>
      <br>

      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr>
          <td width="100%">
            <p align="center"><font color="#FFFFFF"><a href="read.asp?user_id=<%=session("user_id")%>" >�鿴��������</a></font></td>
        </tr>
      </table>

    </td>
    <td width="3%" height="65" valign="top">++</td>
    <td width="50%" height="65" valign="top" style="border-left-style: solid; border-left-width: 1">
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr>
          <td width="100%" bgcolor="#336699"><b><font color="#FFFFFF">������Ƭ�ϴ�</font></b></td>
        </tr>
        <tr>
          <td width="100%">&nbsp;</td>
        </tr>
        <tr>
          <td width="100%"><b>��ע��<span class="content">��Ƭ���</span>:</b></td>
        </tr>
        <tr>
          <td width="100%">1.��Ƭ�ļ���С��60K����</td>
        </tr>
        <tr>
          <td width="100%">2.��Ƭ��ʽ��GIF��JPG</td>
        </tr>
        <tr>
          <td width="100%">&nbsp;</td>
        </tr>
        <tr>
          <td width="100%"><b>�ϴ���Ƭ����˵��</b></td>
        </tr>
        <tr>
          <td width="100%">1.�����ϴ�ɫ��ͼƬ,����Υ��һ�к���Ը�</td>
        </tr>
        <tr>
          <td width="100%">2.�����ϴ��뱾���޹ص���Ƭ</td>
        </tr>
        <tr>
          <td width="100%">&nbsp;&nbsp;</td>
        </tr>
        <tr>
          <td width="100%">����������з����к��������Ƭ���ϣ���</td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<map name="FPMap1"> <area shape="rect" coords="381, 4, 446, 18" href="sendphoto.asp"> 
  <area shape="rect" coords="283, 4, 357, 18" href="register.asp"
 target="_blank"> <area shape="rect" coords="190, 4, 264, 18" href="list.asp"> 
  <area shape="rect" coords="8, 2, 71, 16" href="your.asp"> <area shape="rect" coords="102, 4, 168, 18" href="no.htm"> 
</map>
<map name="Map">
  <area shape="rect" coords="233,95,277,111" href="default.asp">
  <area shape="rect" coords="292,95,358,111" href="your.asp">
  <area shape="rect" coords="381,95,447,118" href="list.asp">
  <area shape="rect" coords="468,94,533,115" href="register.asp">
  <area shape="rect" coords="557,94,621,114" href="sendphoto.asp">
  <area shape="rect" coords="641,93,704,115" href="pics.asp">
</map>
</body>

</html>
</script>






