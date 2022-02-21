<!--#include file="adovbs.inc" -->
<!--#include file="ShowPage.asp" -->
<!--#include file="Functions.asp" -->
<%
	Cata = request("D")
	KeyW = request("StrL")
	Page=request("Page")

	if Cata <> "所有类别" then
		sql="select * from 商品信息 where 名称 like '%" + KeyW + "%' and 分类='" + Cata + "'"
	else
		sql="select * from 商品信息 where 名称 like '%" + KeyW + "%'"
	end if	
	Set conn = OpenOrGet_Database("MyOffers")
%>

<html>
<head>
<title>网上购物网站</title>
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

<script language="javascript" src="java_function.js"></script>

</head>

<body bgcolor="#FFFFFF" text="#000000" onLoad="MM_preloadImages('indeximage/kefu02.gif','indeximage/gouwu02.gif','indeximage/shoucang02.gif','indeximage/shouyin02.gif')" topmargin="2" leftmargin="2">


<div align="center">
  <center>
      <table width="697" border="0" cellpadding="0" cellspacing="0" height="92">
        <tr> 
          <td rowspan="2" width="140" height="92"><img src="indeximage/logo.gif" width="154" height="104"></td>
          <td height="68" colspan="6" width="595"><img src="indeximage/banner.gif" width="611" height="62"></td>
        </tr>
        <tr> 
          <td width="299" height="24"><img src="indeximage/label1.gif" width="299" height="36" usemap="#Map" border="0"></td>
            <td width="76" height="24"><a href style="CURSOR: hand" onClick="window.open('shopbag.asp',null,'width=560,height=360,scrollbars=1')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image40','','indeximage/gouwu02.gif',1)" target="_blank"><img name="Image40" border="0" src="indeximage/gouwu01.gif" width="76" height="36"></a></td>
            <td width="76" height="24"><a href style="CURSOR: hand" onClick="window.open('favorite.asp',null,'width=560,height=360,scrollbars=1')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image37','','indeximage/shoucang02.gif',1)" target="_blank"><img name="Image37" border="0" src="indeximage/shoucang01.gif" width="76" height="36"></a></td>
		    <td width="76" height="24"><a href style="CURSOR: hand" onClick="CheckCookie()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image38','','indeximage/shouyin02.gif',1)" target="_blank"><img name="Image38" border="0" src="indeximage/shouyin01.gif" width="76" height="36"></a></td>
            <td width="84" height="24"><a href="server.htm" style="CURSOR: hand" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image39','','indeximage/kefu02.gif',1)"><img name="Image39" border="0" src="indeximage/kefu01.gif" width="84" height="36"></a></td>
        </tr>
      </table>
  <table width="728" border="0" cellpadding="0" cellspacing="0">
    <tr> 
        <td width="191" align="left" valign="top" height="598" background="indeximage/bottom.gif"> 
          
        <table width="183" border="0" cellpadding="0" cellspacing="0" height="588">
          <tr> 
            <td height="38" width="157" background="indeximage/ditu01.gif" valign="top" align="center"> 
              <form name="Searching" method="POST" action="SearchResult.asp">
                <table width="94%" border="0" cellspacing="0" cellpadding="2">
                  <tr> 
                    <td width="30%"><font size="2"><b>类别:</b></font></td>
                    <td width="70%" colspan="2"> 
                      <select size="1" name="D">
                        <option selected value="所有类别">--所有类别-</option>
                        <option value="长篇评书">长篇评书</option>
                        <option value="传统相声">传统相声</option>
                        <option value="名著赏析">名著赏析</option>
                        <option value="其他内容">其他内容</option>
                        <option value="电脑学习">电脑学习</option>
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
                      </select>
                    </td>
                  </tr>
                  <tr> 
                    <td width="30%" height="22"><font size="2"><b>字串:</b></font></td>
                    <td width="39%" height="22"> 
                      <input type="text" name="StrL" size="7">
                    </td>
                    <td width="31%" height="22" valign="bottom" align="right"><a href style="CURSOR:hand" onClick="
                    	if (Searching.StrL.value == '')
                    		{
                    			alert('您忘记填写查询内容了!')
                    		}
                    	else
                    		{
                    			Searching.submit()
                    		}
                    	"><img src="indeximage/search.gif" width="38" height="19"></a></td>
                  </tr>
                </table>
              </form>
            </td>
            <td height="38" width="34"><img src="indeximage/01.gif" width="34" height="78"></td>
          </tr>
          <tr> 
            <td width="157" height="8"><img src="indeximage/kind01.gif" width="157" height="32" border="0"></td>
            <td width="34" height="8" valign="top"><img src="indeximage/02.gif" width="34" height="32"></td>
          </tr>
          <tr> 
            <td width="157" background="indeximage/ditu03.gif" align="center" height="2" valign="top"> 
              <table width="80%" border="0" cellpadding="2" cellspacing="0">
                <tr> 
                  <td width="46%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?C=长篇评书'">长篇评书</a></font></td>
                  <td width="54%" align="right"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?C=传统相声'">传统相声</a></font></td>
                </tr>
                <tr> 
                  <td width="46%" height="2"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?C=名著赏析'">名著赏析</a></font></td>
                  <td width="54%" align="right" height="2"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?C=长篇评书'">其他内容</a></font></td>
                </tr>
              </table>
            </td>
            <td width="34" background="indeximage/03.gif" valign="top" height="2"><img src="indeximage/03.gif" width="34" height="30"></td>
          </tr>
          <tr> 
            <td width="157" height="20"><img src="indeximage/kind02.gif" width="157" height="35" border="0"></td>
            <td width="34" height="20" valign="top"><img src="indeximage/04.gif" width="34" height="35"></td>
          </tr>
          <tr> 
            <td width="157" background="indeximage/ditu05.gif" align="center" height="60"> 
              <table width="82%" border="0" cellpadding="2" cellspacing="0">
                <tr> 
                  <td width="46%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=电脑学习'">电脑学习</a></font></td>
                  <td width="54%" align="right"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=外语学习'">外语学习</a></font></td>
                </tr>
                <tr> 
                  <td width="46%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=家庭百科'">家庭百科</a></font></td>
                  <td width="54%" align="right"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=多媒体类'">多媒体类</a></font></td>
                </tr>
                <tr> 
                  <td width="46%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=查毒杀毒'">查毒杀毒</a></font></td>
                  <td width="54%" align="right"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=系统应用'">系统应用</a></font></td>
                </tr>
              </table>
            </td>
            <td width="34" height="60" valign="top"><img src="indeximage/05.gif" width="34" height="64"></td>
          </tr>
          <tr> 
            <td width="157" height="16"><img src="indeximage/kind03.gif" width="157" height="36" border="0"></td>
            <td width="34" height="16" valign="top"><img src="indeximage/06.gif" width="34" height="36"></td>
          </tr>
          <tr> 
            <td width="157" background="indeximage/ditu07.gif" align="center" height="35"> 
              <table width="83%" border="0" cellpadding="2" cellspacing="0">
                <tr> 
                  <td width="46%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='game.asp?Catalog=角色扮演'">角色扮演</a></font></td>
                  <td width="54%" align="right"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='game.asp?Catalog=动作冒险'">动作冒险</a></font></td>
                </tr>
                <tr> 
                  <td width="46%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='game.asp?Catalog=战略策略'">战略策略</a></font></td>
                  <td width="54%" align="right"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='game.asp?Catalog=模拟经营'">模拟经营</a></font></td>
                </tr>
                <tr> 
                  <td width="46%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='game.asp?Catalog=益智休闲'">益智休闲</a></font></td>
                  <td width="54%" align="right"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='game.asp?Catalog=体育竞技'">体育竞技</a></font></td>
                </tr>
              </table>
            </td>
            <td width="34" height="35" valign="top"><img src="indeximage/07.gif" width="34" height="58"></td>
          </tr>
          <tr> 
            <td width="157" height="36"><img src="indeximage/kind04.gif" width="157" height="40" border="0"></td>
            <td width="34" height="36" valign="top"><img src="indeximage/08.gif" width="34" height="40"></td>
          </tr>
          <tr> 
            <td width="157" background="indeximage/ditu09.gif" align="center" height="6" valign="top"> 
              <table width="85%" border="0" cellpadding="2" cellspacing="0">
                <tr> 
                  <td width="46%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href=''">MP3播放</a></font></td>
                  <td width="54%" align="right"><font size="2"><a href onClick="this.href=''"> 
                    </a></font></td>
                </tr>
              </table>
            </td>
            <td width="34" height="6" valign="top"><img src="indeximage/09.gif" width="34" height="45"></td>
          </tr>
          <tr> 
            <td width="157" height="2"><img src="indeximage/kind05.gif" width="157" height="23"></td>
            <td width="34" height="2"><img src="indeximage/10.gif" width="34" height="23"></td>
          </tr>
          <tr> 
            <td width="157" height="41" valign="top"> 
              
            </td>
            <td width="34" valign="top" height="41">&nbsp;</td>
          </tr>
          <tr> 
            <td width="157" background="indeximage/ditu12.gif" height="63">&nbsp;</td>
            <td width="34" background="indeximage/ditu13.gif" height="63">&nbsp;</td>
          </tr>
        </table>
        </td>
        
      <td width="542" align="left" valign="top" height="598"> 
        <table width="534" border="0" cellspacing="0" cellpadding="0" height="170">
          <tr> 
            <td height="9" valign="top" width="532">                                                                 
                                        
            <img border="0" src="indeximage/13.gif">                                                                 
                                        
            </td>                            
          </tr> 
          <tr> 
            <td height="22" valign="top" width="532">
            
<%               
	If Request("Page") <> "" Then              
		Set rs = OpenOrGet_RsAndPageSize( conn, sql, "Offer_rs" )               
	Else               
		Set rs = Open_RsAndPageSize( conn, sql, "Offer_rs" )               
	End If
		
  
	Page = CLng(Request("Page"))         
	If Page <1 	Then Page= 1               
	If Page> rs.PageCount Then Page = rs.PageCount               
%>              
              
<form action="SearchResult.asp" method="POST">              
  <div align="right"><p>
<%               
   If Page <> 1 Then              
      Response.Write "<A HREF=SearchResult.asp?Page=1><font size=2 color=blue> 第一页 </font></A>"              
      Response.Write "<A HREF=SearchResult.asp?Page=" & (Page-1) & "><font size=2 color=blue> 上一页 </font> </A>"              
   End If              
   If Page <> rs.PageCount Then              
      Response.Write "<A HREF=SearchResult.asp?Page=" & (Page+1) & "><font size=2 color=blue> 下一页 </font> </A>"              
      Response.Write "<A HREF=SearchResult.asp?Page=" & rs.PageCount & "><font size=2 color=blue> 最后一页 </font> </A>"              
   End If  
   
%>              
    <font size="2">输入<input type="text" size="3" name="Page">                                                                     
    页数</font><font color="#FF0000"><%=Page%>/<%=rs.PageCount%></font> </p>                                                                     
  </div>                                                                     
</form>                                                                     
                                            
            </td>                                
          </tr>                                
          <tr>                                 
            <td height="139" valign="top" width="532">                                 
            <% if page>=0 then                             
            	ShowOnePage rs, Page    
				end if                                                                                                                                 
			%>                                                                                       
            </td>                                                                                       
          </tr>                                                                                       
        </table>                                                                                                                            
    </td>                                                                                                                            
  </tr>                                                                                                                            
</table>                                                                                                                            
</div>                                                                                                                            
<map name="Map">                                                                                                                             
  <area shape="rect" coords="28, 4, 79, 21" href="index.asp">                                                                                                                            
  <area shape="rect" coords="88, 4, 141, 21" href="listen.htm">                                                                                                                            
  <area shape="rect" coords="157, 4, 210, 22" href="play.htm">                                                                                                                            
  <area shape="rect" coords="225, 4, 278, 22" href="xinpin.htm">                                                                                                                            
</map>                                                                                                                            
<div align="center">                                                                                                                         
  <center>                                                                                                                         
    <table width="768" border="0" cellpadding="0" cellspacing="0">                                                                                      
      <tr>                                                                                      
        <td width="191" height="18">&nbsp;</td>                                                                                      
        <td width="577" align="center" height="18">&nbsp;</td>                                                                                      
      </tr>                                                                                      
      <tr>                                                                                       
        <td width="191">&nbsp;</td>                                                                                      
        <td width="577" align="center"><font size="2">copyright:郑州亮中</font></td>                                                                                     
      </tr>                                                                                     
      <tr>                                                                                      
        <td colspan="2">                                                                                      
          <hr width="100%" size="1">                                                                                     
        </td>                                                                                     
      </tr>                                                                                     
      <tr>                                                                                      
        <td width="191">&nbsp;</td>                                                                                     
        <td width="577" align="center"><font size="2">              
          E-mail:<a href="mailto:xxxx@xxxx.com">xxxx@xxxx.com</a></font></td>                                                                                      
      </tr>                                                                                      
      <tr>                                                                                       
        <td width="191">&nbsp;</td>                                                                                      
        <td width="577">&nbsp; </td>                                                                                      
      </tr>                                                                                      
    </table>                                                                                                                          
</center>                                                                                                                         
</div>                                                                                                                          
</body>                                                                                                                            
</html>                                                                                                                            
                                                                                           
                                                                                           
