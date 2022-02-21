<!--#include file="adovbs.inc" -->
<!--#include file="ShowPage.asp" -->
<!--#include file="Functions.asp" -->
<html>
<head>
<title>软件商城</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>

<STYLE>
TD {
	COLOR: #282828; FONT-FAMILY: 宋体; FONT-SIZE: 9pt; LINE-HEIGHT: 16px
}
</STYLE>



</head>

<body bgcolor="#FFFFFF" text="#000000" onLoad="MM_preloadImages('indeximage/gouwu02.gif','indeximage/shoucang02.gif','indeximage/shouyin02.gif','indeximage/kefu02.gif')" topmargin="0" leftmargin="2">
<div align="center"><center> 
<table width="768" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td rowspan="2"><img src="indeximage/logo.gif" width="157" height="98"></td>
    <td height="62" colspan="6"><img src="indeximage/banner.gif" width="611" height="62"></td>
  </tr>
  <tr> 
    <td><img src="indeximage/label1.gif" width="299" height="36" usemap="#Map" border="0"></td>
      <td> <a href style="CURSOR: hand" onClick="window.open('shopbag.asp',null,'width=560,height=360,scrollbars=1')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image40','','indeximage/gouwu02.gif',1)" target="_blank"><img name="Image40" border="0" src="indeximage/gouwu01.gif" width="76" height="36"></a></td>
      <td><a href style="CURSOR: hand" onClick="window.open('favorite.asp',null,'width=560,height=360,scrollbars=1')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image37','','indeximage/shoucang02.gif',1)" target="_blank"><img name="Image37" border="0" src="indeximage/shoucang01.gif" width="76" height="36"></a></td>
      <td><a href style="CURSOR: hand" onClick="window.open('money.asp',null,'width=560,height=360,scrollbars=1')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image38','','indeximage/shouyin02.gif',1)" target="_blank"><img name="Image38" border="0" src="indeximage/shouyin01.gif" width="76" height="36"></a></td>
      <td><a href="server.htm" style="CURSOR: hand" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image39','','indeximage/kefu02.gif',1)"><img name="Image39" border="0" src="indeximage/kefu01.gif" width="84" height="36"></a></td>
  </tr>
</table>
<map name="Map"> 
  <area shape="rect" coords="89, 3, 145, 25" href="listen.htm">
  <area shape="rect" coords="225, 3, 278, 24" href="/xinpin.htm">
  <area shape="rect" coords="159, 3, 212, 24" href="play.htm">
  <area shape="rect" coords="25, 4, 78, 25" href="index.asp">
</map>
<table width="667" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="169" height="388" valign="top" background="image/ditu07.gif"> 
      <table width="169" border="0" cellspacing="0" cellpadding="0" height="519">
        <tr> 
          <td width="34%" valign="top" background="image/ditu01.gif" height="84"> 
            <form name="Searching" method="POST" action="SearchResult.asp">
            <table width="96%" border="0" cellspacing="0" cellpadding="2">
              <tr> 
                <td width="31%"><font size="2"><b>类别:</b></font></td>
                <td width="69%" colspan="2"> 
                  <select size="1" name="D">
                    <option selected value="电脑学习">电脑学习</option>
                    <option value="长篇评书">长篇评书</option>
                    <option value="传统相声">传统相声</option>
                    <option value="名著赏析">名著赏析</option>
                    <option value="其他内容">其他内容</option>
                    <option value="外语学习">外语学习</option>
                    <option value="家庭百科">家庭百科</option>
                    <option value="多媒体类">多媒体类</option>
                    <option value="查毒杀毒">查毒杀毒</option>
                    <option value="系统应用">系统应用</option>
                    <option value="角色扮演">角色扮演</option>
                    <option value="动作冒险">动作冒险</option>
                    <option value="战略策略">战略策略</option>
                    <option value="模拟经营">模拟经营</option>
                    <option value="益智休闲">益智休闲</option>
                    <option value="体育竞技">体育竞技</option>
                    <option value="所有类别">--所有类别-</option>
                  </select>
                </td>
              </tr>
              <tr> 
                <td width="31%"><font size="2"><b>字串:</b></font></td>
                <td width="40%"> 
                  <input type="text" name="StrL" size="7">
                </td>
                <td width="29%"><a href style="CURSOR:hand" onClick="
                    	if (Searching.StrL.value == '')
                    		{
                    			alert('您忘记填写查询内容了!')
                    		}
                    	else
                    		{
                    			Searching.submit()
                    		}
                    	"> 
                  <img src="indeximage/search.gif" width="38" height="19">
                  </a>
                </td>
              </tr>
            </table>
           </form>
          </td>
          <td width="66%" valign="top" height="84"><img src="image/001.gif" width="14" height="78"></td>
        </tr>
        <tr> 
          <td valign="top" width="34%" height="38"><img src="image/soft01.gif" width="155" height="32"></td>
          <td valign="top" width="66%" height="38"><img src="image/002.gif" width="14" height="32"></td>
        </tr>
        <tr align="left" valign="top"> 
          <td width="34%" background="image/ditu02.gif" height="30"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=电脑学习'"><img src="image/soft02.gif" width="155" height="63"></a></td>
          <td width="66%" height="30" background="image/003.gif"><img src="image/soft09.gif" width="14" height="63"></td>
        </tr>
        <tr align="left" valign="top"> 
          <td width="34%" background="image/ditu03.gif" height="53"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=外语学习'"><img src="image/soft03.gif" width="155" height="47"></td>
          <td width="66%" background="image/004.gif" height="53"><img src="image/soft10.gif" width="14" height="47"></td>
        </tr>
        <tr align="left" valign="top"> 
          <td width="34%" background="image/ditu04.gif" height="53"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=家庭百科'"><img src="image/soft04.gif" width="155" height="47"></td>
          <td width="66%" background="image/005.gif" height="53"><img src="image/soft11.gif" width="14" height="47"></td>
        </tr>
        <tr align="left" valign="top"> 
          <td width="34%" background="image/ditu05.gif" height="49"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=多媒体类'"><img src="image/soft05.gif" width="155" height="43"></td>
          <td width="66%" background="image/ditu06.gif" height="49"><img src="image/soft12.gif" width="14" height="43"></td>
        </tr>
        <tr align="left" valign="top"> 
          <td width="34%" background="image/ditu05.gif" height="51"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=查毒杀毒'"><img src="image/soft06.gif" width="155" height="45"></td>
          <td width="66%" background="image/ditu06.gif" height="51"><img src="image/soft13.gif" width="14" height="45"></td>
        </tr>
        <tr align="left" valign="top">
          <td width="34%" height="55"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=系统应用'"><img src="image/soft07.gif" width="155" height="49"></a></td>
          <td width="66%" height="55"><img src="image/soft14.gif" width="14" height="49"></td>
        </tr>
        <tr align="left" valign="top"> 
          <td width="34%" height="106"><img src="image/soft08.gif" width="155" height="100"></td>
          <td width="66%" height="106">&nbsp;</td>
        </tr>
      </table>
    </td>   
      <td width="599" height="426" valign="top"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="2">
        
        <tr> 
            <td height="5"><img src="indeximage/13back.gif" width="596" height="16"></td>
        </tr>
        <tr> 	 

          <td height="5">
           
          
<%               
	SQL="Select * from 商品信息 where 分类='" + Request("Catalog") + "'"
	Set conn = OpenOrGet_Database("MyOffers")

	If Request("Page") <> "" Then              
		Set rs = OpenOrGet_RsAndPageSize( conn, sql, "Offer_rs" )               
	Else               
		Set rs = Open_RsAndPageSize( conn, sql, "Offer_rs" )               
	End If
		
	Page = CLng(Request("Page"))         
	If Page <1 	Then Page= 1               
	If Page> rs.PageCount Then Page = rs.PageCount               
%>              
              
<form action="soft.asp" method="POST">              
  <div align="right"><p>
<%               
   If Page <> 1 Then              
      Response.Write "<A HREF=soft.asp?Page=1><font size=2 color=blue>第一页</font></A>"              
      Response.Write "&nbsp; <A HREF=soft.asp?Page=" & (Page-1) & "><font size=2 color=blue> 上一页</font> </A>"              
   End If              
   If Page <> rs.PageCount Then              
      Response.Write "&nbsp; <A HREF=soft.asp?Page=" & (Page+1) & "><font size=2 color=blue> 下一页</font></A>"              
      Response.Write "&nbsp; <A HREF=soft.asp?Page=" & rs.PageCount & "><font size=2 color=blue>最末页</font></A>&nbsp; "              
   End If  
   
%>              
    <font size="2">输入<input type="text" size="3" name="Page">                                                                       
    页数</font><font color="#FF0000"><%=Page%>/<%=rs.PageCount%></font> </p>                                                                       
  </div>                                                                       
</form>                                                                       
                                                                        
                                                                    
    </td>                                                              
  </tr>                                                           
  <tr>                                 
    <td height="135" valign="top">                                 
    <%                              
       	ShowOnePage rs, Page                                                                                                                                     
	%>                                                                                       
    </td>                                                                                       
  </tr>                                                                                       
           
                                                          
      </table>                                                          
      </td>                                                          
  </tr>                                                          
</table>                          
<table width="768" border="0" cellspacing="0" cellpadding="0">                   
  <tr>                           
    <td colspan="2" width="768">                           
      <hr width="100%" size="1" noshade>                          
    </td>                          
  </tr>                          
  <tr>                           
    <td width="168" height="4">&nbsp;</td>                          
    <td width="606" height="4" align="center"><font size="2"><b>Copyright</b>:郑州亮中</font></td>                          
  </tr>                          
  <tr>                           
    <td width="168" height="4">&nbsp;</td>                          
    <td width="606" height="4" align="center"><font size="2">　E-mail:<a href="mailto:xxxx@xxxx.com">         
    	xxxx@xxxx.com</a></font></td>                         
  </tr>                         
  <tr>                          
    <td width="168" height="6">&nbsp;</td>                         
    <td width="606" height="6">&nbsp;</td>                         
  </tr>                         
</table>                         
</body>                         
</html>                     