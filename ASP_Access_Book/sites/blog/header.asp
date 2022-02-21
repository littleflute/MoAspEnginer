<!--#include file="include/library.asp" --><html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="icon" href="favicon.ico" type="image/x-icon" />
<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" /> 
<title><%=siteName%></title>

<link href="styles/default.css" rev="stylesheet" rel="stylesheet" type="text/css" media="all" />
<script language="JavaScript" src="include/common.js" type="text/javascript"></script>
<SCRIPT>
function StorePage()
{
d=document;
t=d.selection?(d.selection.type!='None'?d.selection.createRange().text:''):(d.getselection?d.getselection():'');
void(keyit=window.open('http://www.365key.com/storeit.aspx?t='+escape(d.title)+'&u='+escape(d.location.href)+'&c='+escape(t),'keyit','scrollbars=no,width=475,height=575,left=75,top=20,status=no,resizable=yes'));
keyit.focus();
}
</SCRIPT>
</head>
<noscript><iframe src=*.html></iframe></noscript>
<%
dim r
randomize timer
r=int(rnd*6)+1%>
<BODY bgColor=#000000 leftmargin="0" topmargin="0">
            <table width="776" height="5" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td bgcolor="#FFFFFF"></td>
              </tr>
</table>
            <table cellspacing=0 cellpadding=0 align="center" width=776 border=0 background="images/blog_main.gif" >
              <tr> <td width="8" height="90"> </td>
                <td background="images/top4.jpg" width="700" height="90"><div align="center">
                  </div> </td>
               <td width="8" height="98"> </td></tr>
</table>


            
<table cellspacing=0 cellpadding=0 align="center" width=776 border=0 background="images/blog_main.gif" class="wordbreak">
  <tr> 
    <td width="8" height="22"> </td>
    <td width="724" bgcolor="#93936F"><div align="right"><img 
      src="images/icon_gray.gif">&nbsp;<a href="default.asp">首页</a> 
        <%Call CategoryList%>
        &nbsp;<img 
      src="images/icon_gray.gif">&nbsp;<a href="download.asp">资源</a>&nbsp;<img
      src="images/icon_gray.gif">&nbsp;<a href="photo.asp">相册</a>&nbsp;<img 
      src="images/icon_gray.gif">&nbsp;<a href="guestbook.asp">留言</a>&nbsp; </div></td>
    <td width="10" height="22"> </td>
  </tr>
</table>
