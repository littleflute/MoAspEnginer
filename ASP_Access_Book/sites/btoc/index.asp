<%
	response.cookies("times")=0
%>

<html>
<head>
<title>洋洋购物商城欢迎您！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="work/纪录/2002.4.23/work/网站/style.css">
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

<body bgcolor="#006699" text="#000000" onLoad="MM_preloadImages('indeximage/gouwu02.gif','indeximage/shoucang02.gif','indeximage/kefu02.gif','indeximage/shouyin02.gif')" topmargin="0" leftmargin="2">
<div align="center">
  <center>
  <table width="768" border="0" cellpadding="0" cellspacing="0">
    <tr>
    <td background="/indeximage/ditu.gif">
        <table width="744" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td rowspan="2" width="157"><img src="/indeximage/logo.gif" width="195" height="98"></td>
            <td height="62" colspan="5"><img border="0" src="indeximage/banner.gif" width="611" height="62"></td>
          </tr>
          <tr> 
            <td width="299"><img src="/indeximage/label1.gif" usemap="#Map" border="0" width="299" height="36"></td>
      		<td width="76"> <a href style="CURSOR: hand" onClick="window.open('shopbag.asp',null,'width=560,height=360,scrollbars=1')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image40','','indeximage/gouwu02.gif',1)" target="_blank"><img name="Image40" border="0" src="/indeximage/gouwu01.gif" width="76" height="36"></a></td>
      		<td width="76"><a href style="CURSOR: hand" onClick="window.open('favorite.asp',null,'width=560,height=360,scrollbars=1')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image37','','indeximage/shoucang02.gif',1)" target="_blank"><img name="Image37" border="0" src="/indeximage/shoucang01.gif" width="76" height="36"></a></td>
      		<td width="76"><a href style="CURSOR: hand" onClick="window.open('money.asp',null,'width=560,height=360,scrollbars=1')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image38','','indeximage/shouyin02.gif',1)" target="_blank"><img name="Image38" border="0" src="/indeximage/shouyin01.gif" width="76" height="36"></a></td>
      		<td width="84"><a href="server.htm" style="CURSOR: hand" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image39','','indeximage/kefu02.gif',1)"><img name="Image39" border="0" src="/indeximage/kefu01.gif" width="84" height="36"></a></td>
          </tr>
        </table>
    </td>
  </tr>
</center>                                                              
</table> 
  <center>
  <table width="741" border="0" cellpadding="0" cellspacing="0"> 
    <tr>  
        
      <td width="191" align="left" valign="top" height="598">
<TABLE cellSpacing=0 cellPadding=0 width=173 align=center bgColor=#ffffff 
border=0>
          <TR> 
            <TD width=174 height="222" align=middle vAlign=top bgColor=#dedede> 
              <TABLE cellSpacing=0 cellPadding=0 width=172 border=0>
                <TBODY>
                  <TR> 
                    <TD height=4></TD>
                  </TR>
                  <TR> 
                    <TD class="box172 b" 
          background="z/index_pcedu_14.gif" 
          height=23><table width="88%" height="29" border="0">
                        <tr>
                          <td width="17%">&nbsp;</td>
                          
                        <td width="83%"><font color="#FFFFFF"><strong>商品搜索</strong></font></td>
                        </tr>
                      </table> </TD>
                  </TR>
                  <TR> 
                    <TD height=4></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <TABLE cellSpacing=1 cellPadding=0 width=164 bgColor=#c6c6c6 border=0>
                <TBODY>
                  <TR> 
                    
                  <TD vAlign=top align=middle bgColor=#ffffff height=79> 
                    <table cellspacing=0 cellpadding=3 width=160 align=center 
              border=0>
                      <form name="Searching" method="POST" action="SearchResult.asp">  <tr align=middle> 
                        <td height=40><font size="2"><b>类别:</b></font> </td>
                        <td> 
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
                      <tr align=middle> 
                        <td height=40 style="LINE-HEIGHT: 20px"> 
                            <div align="center"><font size="2"><b>字串:</b></font></div>
                        </td>
                        <td> 
                          <input type="text" name="StrL" size="12">
                        </td>
                      </tr>
                      <tr align=middle> 
                        <td height=40 colspan="2"> 
                            <div align="center"> <a href style="CURSOR:hand" onClick="
                    	if (Searching.StrL.value == '')
                    		{
                    			alert('您忘记填写查询内容了!')
                    		}
                    	else
                    		{
                    			Searching.submit()
                    		}
                    	"><img height=19 hspace=5 width=50 
                  src="z/2005_soft_14.gif" 
                  vspace=2 name=image32 border="0"></a> </div>
                        </td>
                      </tr>
                      <tbody> </form>
                    </table>
                  </TD>
                  </TR>
              </TABLE>
              <TABLE cellSpacing=0 cellPadding=0 width=172 border=0>
                <TBODY>
                  <TR> 
                    <TD height=4></TD>
                  </TR>
                  <TR> 
                    <TD class="box172 b" 
          background="z/index_pcedu_14.gif" 
          height=23><table width="88%" height="29" border="0">
                        <tr> 
                          <td width="17%">&nbsp;</td>
                          
                        <td width="83%"><font color="#FFFFFF"><strong>家佳听书馆</strong></font></td>
                        </tr>
                      </table></TD>
                  </TR>
                  <TR> 
                    <TD height=4></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <TABLE cellSpacing=1 cellPadding=0 width=164 bgColor=#c6c6c6 border=0>
                <TBODY>
                  <TR> 
                    <TD vAlign=top align=middle bgColor=#f0f5f9 > 
                    <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                      <TBODY> 
                      <TR> 
                        <TD style="LINE-HEIGHT: 20px"><FONT color=#ff9900>·</FONT><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?C=长篇评书'">长篇评书</a></font> 
                          <BR>
                          <FONT 
                  color=#ff9900>·</FONT><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?C=传统相声'">传统相声</a></font><BR>
                          <FONT 
                  color=#ff9900>·</FONT><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?C=名著赏析'">名著赏析</a></font> 
                          <BR>
                          <FONT 
                  color=#ff9900>·</FONT><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?C=长篇评书'">其他内容</a></font></TD>
                      </TR>
                      </TBODY> 
                    </TABLE>
                      <TABLE cellSpacing=0 cellPadding=0 width="97%" border=0>
                        <TBODY>
                         
                        </TBODY>
                      </TABLE></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <TABLE cellSpacing=0 cellPadding=0 width=172 border=0>
                <TBODY>
                  <TR> 
                    <TD height=4></TD>
                  </TR>
                  <TR> 
                    <TD class="box172 b" 
          background="z/index_pcedu_14.gif" 
          height=23><table width="88%" height="29" border="0">
                        <tr> 
                          <td width="17%">&nbsp;</td>
                          
                        <td width="83%"><font color="#FFFFFF"><strong>软件商城</strong></font></td>
                        </tr>
                      </table></TD>
                  </TR>
                  <TR> 
                    <TD height=4></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <TABLE cellSpacing=1 cellPadding=0 width=164 bgColor=#c6c6c6 border=0>
                <TBODY>
                  <TR> 
                    
                  <TD vAlign=top align=middle bgColor=#ffffff> 
                    <TABLE cellSpacing=1 cellPadding=0 width=162 align=center 
              border=0>
                      <TBODY> 
                      <TR> 
                        <TD align=middle width=5 height=18>&nbsp;</TD>
                        <TD vAlign=bottom width=68><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=电脑学习'">电脑学习</a></font></TD>
                        <TD vAlign=bottom width=5 bgColor=#ffcc77></TD>
                        <TD align=middle width=5>&nbsp;</TD>
                        <TD vAlign=bottom width=73><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=外语学习'">外语学习</a></font></TD>
                      </TR>
                      <TR> 
                        <TD align=middle bgColor=#ffcc77 colSpan=5 height=1></TD>
                      </TR>
                      <TR> 
                        <TD align=middle width=5 height=18>&nbsp;</TD>
                        <TD vAlign=bottom width=68><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=家庭百科'">家庭百科</a></font></TD>
                        <TD vAlign=bottom width=5 bgColor=#ffcc77></TD>
                        <TD align=middle width=5>&nbsp;</TD>
                        <TD vAlign=bottom width=73><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=多媒体类'">多媒体类</a></font></TD>
                      </TR>
                      <TR> 
                        <TD align=middle bgColor=#ffcc77 colSpan=5 height=1></TD>
                      </TR>
                      <TR> 
                        <TD align=middle width=5 height=18>&nbsp;</TD>
                        <TD vAlign=bottom width=68><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=查毒杀毒'">查毒杀毒</a></font></TD>
                        <TD vAlign=bottom width=5 bgColor=#ffcc77></TD>
                        <TD align=middle width=5>&nbsp;</TD>
                        <TD vAlign=bottom width=73><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='soft.asp?Catalog=系统应用'">系统应用</a></font></TD>
                      </TR>
                      <TR> 
                        <TD align=middle bgColor=#ffcc77 colSpan=5 height=1></TD>
                      </TR>
                      <TR> 
                        <TD align=middle bgColor=#ffcc77 colSpan=5 height=1></TD>
                      </TR>
                      <TR> 
                        <TD align=middle bgColor=#ffcc77 colSpan=5 height=1></TD>
                      </TR>
                      </TBODY> 
                    </TABLE>
                  </TD>
                  </TR>
                </TBODY>
              </TABLE>
              <TABLE cellSpacing=0 cellPadding=0 width=172 border=0>
                <TBODY>
                  <TR> 
                    <TD height=4></TD>
                  </TR>
                  <TR> 
                    <TD class="box172 b" 
          background="z/index_pcedu_14.gif" 
          height=23><table width="88%" height="29" border="0">
                        <tr> 
                          <td width="17%">&nbsp;</td>
                          
                        <td width="83%"><font color="#FFFFFF"><strong>游戏商城</strong></font></td>
                        </tr>
                      </table></TD>
                  </TR>
                  <TR> 
                    <TD height=4></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <TABLE cellSpacing=1 cellPadding=0 width=164 bgColor=#c6c6c6 border=0>
                <TBODY>
                  <TR> 
                    <TD vAlign=top align=middle bgColor=#ffffff height=165> <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                        <TBODY>
                          <TR> 
                            <TD class=time style="LINE-HEIGHT: 20px">
                          <table width="99%" border="0" cellpadding="2" cellspacing="0" align="center">
                            <tr> 
                              <td width="46%"> <font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='game.asp?Catalog=角色扮演'">·角色扮演</a></font></td>
                              <td width="54%" align="right"> 
                                <div align="center"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='game.asp?Catalog=动作冒险'">·动作冒险</a></font></div>
                              </td>
                            </tr>
                            <tr> 
                              <td width="46%"> <font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='game.asp?Catalog=战略策略'">·战略策略</a></font></td>
                              <td width="54%" align="right"> 
                                <div align="center"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='game.asp?Catalog=模拟经营'">·模拟经营</a></font></div>
                              </td>
                            </tr>
                            <tr> 
                              <td width="46%"> <font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='game.asp?Catalog=益智休闲'">·益智休闲</a></font></td>
                              <td width="54%" align="right"> 
                                <div align="center"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='game.asp?Catalog=体育竞技'">·体育竞技</a></font></div>
                              </td>
                            </tr>
                          </table>
                          <BR>
                        </TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                    <table cellspacing=0 cellpadding=0 width=172 border=0>
                      <tbody> 
                      <tr> 
                        <td height=4></td>
                      </tr>
                      <tr> 
                        <td class="box172 b" 
          background="z/index_pcedu_14.gif" 
          height=23>
                          <table width="88%" height="29" border="0">
                            <tr> 
                              <td width="17%">&nbsp;</td>
                              <td width="83%"><font color="#FFFFFF"><strong>网上专卖</strong></font></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                      <tr> 
                        <td height=4></td>
                      </tr>
                      </tbody> 
                    </table>
                    <table cellspacing=0 cellpadding=0 width="100%" border=0>
                      <tbody> 
                      <tr> 
                        <td class=time style="LINE-HEIGHT: 20px"> <font size="2" color="0000ff">·<a href >MP3播放机专卖</a></font><br>
                        </td>
                      </tr>
                      </tbody> 
                    </table>
                    <div align="center"><a href="diaocha.htm"><img border="0" src="image/3.gif"></a></div>
                  </TD>
                  </TR>
                </TBODY>
              </TABLE>
              <table cellspacing=0 cellpadding=0 width=172 border=0>
                <tbody> 
                <tr> 
                  <td height=4></td>
                </tr>
                <tr> 
                  <td class="box172 b" 
          background="z/index_pcedu_14.gif" 
          height=23> 
                    <table width="88%" height="29" border="0">
                      <tr> 
                        <td width="17%">&nbsp;</td>
                        <td width="83%"><font color="#FFFFFF"><strong>洋洋动态</strong></font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height=4></td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 cellpadding=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td bgcolor="#FFFFFF" > <marquee id=board direction='up' scrolldelay='30' scrollamount='1' height=100> 
                    &nbsp; <a href="xinpin.htm" target="_blank"><font color="#FF0000">由于产品计划调整 
                    《隋唐演义》 推迟上市日期待定 </font></a> &nbsp; <a href="xinpin.htm" target="_blank"><font color="#0000FF">《大唐侠女》、《喋血魔窟》、《秘密列车》热卖中 
                    </font></a></marquee> </td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 cellpadding=0 width=172 border=0>
                <tbody> 
                <tr> 
                  <td height=4></td>
                </tr>
                <tr> 
                  <td class="box172 b" 
          background="z/index_pcedu_14.gif" 
          height=23> 
                    <table width="88%" height="29" border="0">
                      <tr> 
                        <td width="17%">&nbsp;</td>
                        <td width="83%"><font color="#FFFFFF">----销售排行----</font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height=4></td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 cellpadding=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td class=time style="LINE-HEIGHT: 20px" bgcolor="#FFFFFF"> 
                    &nbsp; ·MCSE认证实录<br>
                    &nbsp; ·评书《三国演义》<br>
                    &nbsp; ·瑞星2002<br>
                    &nbsp; ·马三立相声全集<br>
                    &nbsp; ·评书《岳飞传》<br>
                    &nbsp; ·大众小游戏<br>
                    &nbsp; ·世界名著MP3<br>
                    &nbsp; ·新概念Office六合一教程<br>
                    &nbsp; ·新概念Photoshop 6教程<br>
                    &nbsp; ·冒险游戏真精彩<br>
                    &nbsp; ·新概念Flash教程<br>
                    &nbsp; ·新概念Flash教程<br>
                    <br>
                    &nbsp; 更多新品，敬请期待 </td>
                </tr>
                </tbody> 
              </table>
              <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                <TBODY>
                
                </TBODY>
              </TABLE></TD>
            <TD width=10>&nbsp;</TD>
          </TR>
        </TABLE> 
        
      </td>                                                  
                                                          
      <td width="550" align="left" valign="top" height="598">                                                   
        <table width="82%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" background="topBar_bg.gif">
          <tr>                                                   
            <td height="13" valign="top">                                                   
              <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
                <tr>                                                   
                  <td valign="bottom" width="39%"><img src="/indeximage/kind06.gif" width="232" height="48"></td>                                                  
                  <td valign="bottom" align="left" width="61%"><img src="/indeximage/kind06-1.gif" width="341" height="15"></td>                                                  
                </tr>                                                  
              </table>                                                  
            </td>                                                  
          </tr>                                                  
          <tr>                                                   
            <td height="13" valign="top">&nbsp; </td>                                                  
          </tr>                                                  
          <tr>                                                   
            <td height="139" valign="top">                                                   
            <!--#include file="conn.asp"-->
			  <%                                                  
                                            
	SQL="Select top 4 * from 商品信息 Where 标记='最新' Order by 时间 DESC"                                                  
	set rs=Conn.execute(SQL)                                                  
	Do while not rs.EOF                                                  
		Ostr="window.open('introduce.asp?P=" + rs.fields("名称") + "&T=" + rs.fields("图标") + "',null,'width=520,height=400,scrollbars=1')"                                                  
		Mstr="window.open('shopbag.asp?P=" + rs.fields("名称") + "&PP=" + rs.fields("定价") + "&Price=" + rs.fields("售价") + "',null,'width=560,height=360,scrollbars=1,status=0')"                                                  
		Lstr="window.open('favorite.asp?P=" + rs.fields("名称") + "&Price=" + rs.fields("售价") + "&T=" + rs.fields("图标") + "',null,'width=560,height=360,scrollbars=1')"                                                                               
%>                                                  
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><table width="55%" border="0">
                      <tr>
                        <td rowspan="4" valign="top" width="16%"><a href onClick="<%=Ostr%>"><img src="../pro-image/<%=rs.fields("图标")%>" alt="详细资料" width="70" height="90" style="CURSOR: hand"></a><a href onClick="window.open('introduce.asp?P=站内测试标题&T=C22.jpg ',null,'width=520,height=400,scrollbars=1')"></a></td>
                        <td width="84%"><font size="2" color="#660000"><b><font color="#000066">商品名:</font><font size="2" color="#990000"><b><font color="#000066"><%=rs.fields("名称")%></font></b></font><font color="#000066">(</font><font size="2" color="#990000"><b><font color="#000066"><%=rs.fields("包装")%></font></b></font><font color="#000066">)</font></b></font></td>
                      </tr>
                      <tr>
                        <td width="84%"><font size="2" color="#660000"><u>原始价格:</u></font><font size="2"><u><font color="#660000"><strike><%=rs.fields("定价")%></strike></font></u></font><font size="2" color="#660000"><u>元                                                                                                                                            
                    会员价格:</u></font><font size="2"><u><font color="#660000"><%=rs.fields("售价")%></font></u></font><font size="2" color="#660000"><u>元</u></font></td>
                      </tr>
                      <tr>
                        <td width="84%"><br>
                        <font size="2"><%=rs.fields("简介")%></font></td>
                      </tr>
                      <tr>
                        <td width="84%"><p align="center"><font size="2"><a href style="CURSOR: hand" onClick="<%=Ostr%>"><font color="#0000FF">详细资料</font></a> <font color="#0000FF"><a style="cursor: hand" href onClick="<%=Mstr%>"><img border="0" src="/image/cart.gif">购买</a></font> <font color="#0000FF"><a href style="CURSOR: hand" onClick="<%=Lstr%>"><img border="0" src="/image/favor.gif">收藏</a></font></font></p></td>
                      </tr>
                    <td width="16%"><center>
                    <tr>
                      <td colspan="2"><img src="/indeximage/line.gif" width="280" height="10"></td>
                    </tr>
                  </table></td>
                  <td>
				  <%rs.movenext
				  if not rs.eof then
				  Ostr="window.open('introduce.asp?P=" + rs.fields("名称") + "&T=" + rs.fields("图标") + "',null,'width=520,height=400,scrollbars=1')"                                                  
		Mstr="window.open('shopbag.asp?P=" + rs.fields("名称") + "&PP=" + rs.fields("定价") + "&Price=" + rs.fields("售价") + "',null,'width=560,height=360,scrollbars=1,status=0')"                                                  
		Lstr="window.open('favorite.asp?P=" + rs.fields("名称") + "&Price=" + rs.fields("售价") + "&T=" + rs.fields("图标") + "',null,'width=560,height=360,scrollbars=1')"
		%>
				  <table width="55%" border="0">
                      <tr>
                        <td rowspan="4" valign="top" width="16%"><a href onClick="<%=Ostr%>"><img src="../pro-image/<%=rs.fields("图标")%>" alt="详细资料" width="70" height="90" style="CURSOR: hand"></a><a href onClick="window.open('introduce.asp?P=站内测试标题&T=C22.jpg ',null,'width=520,height=400,scrollbars=1')"></a></td>
                        <td width="84%"><font size="2" color="#660000"><b><font color="#000066">商品名:</font><font size="2" color="#990000"><b><font color="#000066"><%=rs.fields("名称")%></font></b></font><font color="#000066">(</font><font size="2" color="#990000"><b><font color="#000066"><%=rs.fields("包装")%></font></b></font><font color="#000066">)</font></b></font></td>
                      </tr>
                      <tr>
                        <td width="84%"><font size="2" color="#660000"><u>原始价格:</u></font><font size="2"><u><font color="#660000"><strike><%=rs.fields("定价")%></strike></font></u></font><font size="2" color="#660000"><u>元                                                                                                                                            
                    会员价格:</u></font><font size="2"><u><font color="#660000"><%=rs.fields("售价")%></font></u></font><font size="2" color="#660000"><u>元</u></font></td>
                      </tr>
                      <tr>
                        <td width="84%"><br>
                        <font size="2"><%=rs.fields("简介")%></font></td>
                      </tr>
                      <tr>
                        <td width="84%"><p align="center"><font size="2"><a href style="CURSOR: hand" onClick="<%=Ostr%>"><font color="#0000FF">详细资料</font></a> <font color="#0000FF"><a style="cursor: hand" href onClick="<%=Mstr%>"><img border="0" src="/image/cart.gif">购买</a></font> <font color="#0000FF"><a href style="CURSOR: hand" onClick="<%=Lstr%>"><img border="0" src="/image/favor.gif">收藏</a></font></font></p></td>
                      </tr>
                    <td width="16%"><center>
                    <tr>
                      <td colspan="2"><img src="/indeximage/line.gif" width="280" height="10"></td>
                    </tr>
                  </table
				  ><%rs.movenext
				  end if%></td>
                </tr>
              </table>
              <%                                                                                                                                                                                        
	                                                                                                                                                                                        
	loop                                                                                                                                                                                        
	Conn.close                                                                                                                                                                                        
%>                                                                                                                                          
            </td>                                                                                                                                          
          </tr>                                                                                                                                          
          <tr>                                                                                                                                           
            <td height="10">                                                                                                                                           
              <table width="100%" border="0" cellspacing="0" cellpadding="0">                                                                                                                                          
                <tr>                                                                                                                                           
                  <td valign="bottom" width="39%"><img src="/indeximage/kind07.gif" width="232" height="48"></td>                                                                                                                                          
                  <td valign="bottom" align="left" width="61%"><img src="/indeximage/kind06-1.gif" width="341" height="15"></td>                                                                                                                                          
                </tr>                                                                                                                                          
              </table>                                                                                                                                          
            </td>                                                                                                                                          
          </tr>                                                                                                                                          
          <tr>                                                                                                                                           
            <td height="10">&nbsp; </td>                                                                                                                                          
          </tr>                                                                                                                                          
          <tr>                                                                                                                                           
            <td height="139" valign="top">                                                                                                                                           
                 <!--#include file="conn.asp"-->
			  <%                                                                                                                                                                               
                                                                                                                                                                           
	SQL="Select  top 6 * from 商品信息 Where 标记='推荐' Order by id DESC"                                                                                                                                                                               
	set rs=Conn.execute(SQL)                                                                                                                                                                               
	Do while not rs.EOF                                                                                                                                                                               
		Ostr="window.open('introduce.asp?P=" + rs.fields("名称") + "&T=" + rs.fields("图标") + "',null,'width=520,height=400,scrollbars=1')"                                                                                             
		Mstr="window.open('shopbag.asp?P=" + rs.fields("名称") + "&Price=" + rs.fields("售价") + "',null,'width=560,height=360,scrollbars=1')"                                                                                                                                                                               
		Lstr="window.open('favorite.asp?P=" + rs.fields("名称") + "&Price=" + rs.fields("售价") + "&T=" + rs.fields("图标") + "',null,'width=560,height=360,scrollbars=1')"                                                                                                                                                                               
%>                                                                                                                                          
              <table width="100%" border="0" cellspacing="0" cellpadding="0" height="74">                                                                                                                                          
                <tr>                                                                                                 <% if not rs.eof then          %>                               
                  <td height="74" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><table width="55%" border="0">
                          <tr>
                            <td rowspan="4" valign="top" width="16%"><a href onClick="<%=Ostr%>"><img src="../pro-image/<%=rs.fields("图标")%>" alt="详细资料" width="70" height="90" style="CURSOR: hand"></a><a href onClick="window.open('introduce.asp?P=站内测试标题&T=C22.jpg ',null,'width=520,height=400,scrollbars=1')"></a></td>
                            <td width="84%"><font size="2" color="#660000"><b><font color="#000066">商品名:</font><font size="2" color="#990000"><b><font color="#000066"><%=rs.fields("名称")%></font></b></font><font color="#000066">(</font><font size="2" color="#990000"><b><font color="#000066"><%=rs.fields("包装")%></font></b></font><font color="#000066">)</font></b></font></td>
                          </tr>
                          <tr>
                            <td width="84%"><font size="2" color="#660000"><u>原始价格:</u></font><font size="2"><u><font color="#660000"><strike><%=rs.fields("定价")%></strike></font></u></font><font size="2" color="#660000"><u>元                                                                                                                                            
                              会员价格:</u></font><font size="2"><u><font color="#660000"><%=rs.fields("售价")%></font></u></font><font size="2" color="#660000"><u>元</u></font></td>
                          </tr>
                          <tr>
                            <td width="84%"><br>
                                <font size="2"><%=rs.fields("简介")%></font></td>
                          </tr>
                          <tr>
                            <td width="84%"><p align="center"><font size="2"><a href style="CURSOR: hand" onClick="<%=Ostr%>"><font color="#0000FF">详细资料</font></a> <font color="#0000FF"><a style="cursor: hand" href onClick="<%=Mstr%>"><img border="0" src="/image/cart.gif">购买</a></font> <font color="#0000FF"><a href style="CURSOR: hand" onClick="<%=Lstr%>"><img border="0" src="/image/favor.gif">收藏</a></font></font></p></td>
                          </tr>
                        <td width="16%"><center>
                        <tr>
                          <td colspan="2"><img src="/indeximage/line.gif" width="280" height="10"></td>
                        </tr>
                      </table></td>
                      <td><%rs.movenext
					  end if
				  if not rs.eof then
				  Ostr="window.open('introduce.asp?P=" + rs.fields("名称") + "&T=" + rs.fields("图标") + "',null,'width=520,height=400,scrollbars=1')"                                                  
		Mstr="window.open('shopbag.asp?P=" + rs.fields("名称") + "&PP=" + rs.fields("定价") + "&Price=" + rs.fields("售价") + "',null,'width=560,height=360,scrollbars=1,status=0')"                                                  
		Lstr="window.open('favorite.asp?P=" + rs.fields("名称") + "&Price=" + rs.fields("售价") + "&T=" + rs.fields("图标") + "',null,'width=560,height=360,scrollbars=1')"
		%>
                          <table width="55%" border="0">
                            <tr>
                              <td rowspan="4" valign="top" width="16%"><a href onClick="<%=Ostr%>"><img src="../pro-image/<%=rs.fields("图标")%>" alt="详细资料" width="70" height="90" style="CURSOR: hand"></a><a href onClick="window.open('introduce.asp?P=站内测试标题&T=C22.jpg ',null,'width=520,height=400,scrollbars=1')"></a></td>
                              <td width="84%"><font size="2" color="#660000"><b><font color="#000066">商品名:</font><font size="2" color="#990000"><b><font color="#000066"><%=rs.fields("名称")%></font></b></font><font color="#000066">(</font><font size="2" color="#990000"><b><font color="#000066"><%=rs.fields("包装")%></font></b></font><font color="#000066">)</font></b></font></td>
                            </tr>
                            <tr>
                              <td width="84%"><font size="2" color="#660000"><u>原始价格:</u></font><font size="2"><u><font color="#660000"><strike><%=rs.fields("定价")%></strike></font></u></font><font size="2" color="#660000"><u>元                                                                                                                                            
                                会员价格:</u></font><font size="2"><u><font color="#660000"><%=rs.fields("售价")%></font></u></font><font size="2" color="#660000"><u>元</u></font></td>
                            </tr>
                            <tr>
                              <td width="84%"><br>
                                  <font size="2"><%=rs.fields("简介")%></font></td>
                            </tr>
                            <tr>
                              <td width="84%"><p align="center"><font size="2"><a href style="CURSOR: hand" onClick="<%=Ostr%>"><font color="#0000FF">详细资料</font></a> <font color="#0000FF"><a style="cursor: hand" href onClick="<%=Mstr%>"><img border="0" src="/image/cart.gif">购买</a></font> <font color="#0000FF"><a href style="CURSOR: hand" onClick="<%=Lstr%>"><img border="0" src="/image/favor.gif">收藏</a></font></font></p></td>
                            </tr>
                            <td width="16%"><center>
                            <tr>
                              <td colspan="2"><img src="/indeximage/line.gif" width="280" height="10"></td>
                            </tr>
                          </table
				  >
                        <%rs.movenext
				  end if%></td>
                    </tr>
                  </table>
                    </center>                                                                                                                                          
                    </td>                                                                                                                                          
                </tr>                                                                                                                                          
              </table>                                                                                                                                          
              <%                                                                                                                                                                                         
                                                                                                                                                                                        
	loop                                                                                                                                                                                         
	Conn.close                                                                                                                                                                                         
%>                                                                                                                                          
            </td>                                                                                                                                          
          </tr>                                                                                                                                          
          <tr>                                                                                                                                          
            <td height="7">                                                                                                                                          
              <table width="100%" border="0" cellspacing="0" cellpadding="0">                                                                                                                                          
                <tr>                                                                                                                                           
                  <td valign="bottom" width="39%"><img src="/indeximage/kind08.gif" width="232" height="48"></td>                                                                                                                                          
                  <td valign="bottom" align="left" width="61%"><img src="/indeximage/kind06-1.gif" width="341" height="15"></td>                                                                                                                                          
                </tr>                                                                                                                                          
              </table>                                                                                                                                          
            </td>                                                                                                                                          
          </tr>                                                                                                                                          
          <tr>                                                                                                                                           
            <td height="7">&nbsp; </td>                                                                                                                                          
          </tr>                                                                                                                                          
          <tr>                                                                                                                                           
            <td height="89" valign="top">                                                                                                                                           
             <!--#include file="conn.asp"-->
			  <%                                                                                                                                                                                
	                                                                                                                                                                          
	SQL="Select * from 商品信息 Where 标记='清仓' Order by 时间 DESC"                                                                                                                                                
	set rs=Conn.execute(SQL)                                                                                                                                                                                
	Do while not rs.EOF                                                                                                                                                                                
		Ostr="window.open('introduce.asp?P=" + rs.fields("名称") + "&T=" + rs.fields("图标") + "',null,'width=520,height=400,scrollbars=1')"                                                                                             
		Mstr="window.open('shopbag.asp?P=" + rs.fields("名称") + "&Price=" + rs.fields("售价") + "',null,'width=560,height=360,scrollbars=1')"                                                                                                                                                                                
		Lstr="window.open('favorite.asp?P=" + rs.fields("名称") + "&Price=" + rs.fields("售价") + "&T=" + rs.fields("图标") + "',null,'width=560,height=360,scrollbars=1')"                                                                                                                                                                                
%>                                                                                                                                          
              <table width="100%" border="0" cellspacing="0" cellpadding="0" height="74">                                                                                                                                          
                <tr>                                                                                                                                           
                  <td height="74" valign="top">                                                                                                                                           
                    <table width="100%" border="0">                                                                                                                                          
                      <tr>                                                                                                                                           
                        <td rowspan="4" valign="top" width="16%"><a href onClick="<%=Ostr%>"><img src="pro-image/<%=rs.fields("图标")%>" alt="详细资料" width="70" height="90" style="CURSOR: hand" border="0"></a></td>                                                                                                                                          
                        <td width="82%"><font size="2"><b><font color="#000066">商品名:<%=rs.fields("名称")%>                                                                                                                                           
                          (<%=rs.fields("包装")%>)</font></b></font></td>                                                                                                                                           
                      </tr>                                                                                                                                           
                      <tr>                                                                                                                                            
                        <td width="82%"><font size="2" color="#660000"><u>原始价格:<strike><%=rs.fields("定价")%></strike>元                                                                                                                                            
                          会员价格:<%=rs.fields("售价")%>元</u></font></td>                                                                                                                                           
                      </tr>                                                                                                                                           
                      <tr>                                                                                                                                            
                        <td width="82%"><br><font size="2"><%=rs.fields("简介")%></font></td>                                                                                                                                           
                      </tr>                                                                                                                                           
                      <tr>                                                                                                                                            
                        <td width="82%">                                                                                                                                            
                          <p align="right"><font size="2">                                                                                                     
                      <a href style="CURSOR: hand" onClick="<%=Ostr%>"><font color="#0000FF">详细资料</font></a><font color="#0000FF"></font><font color="0000ff"><a style="cursor: hand" href onClick="<%=Mstr%>"><img border="0" src="/image/cart.gif">购买</a></font>                                                                 
                          <a href style="CURSOR: hand" onClick="<%=Lstr%>"><font color="0000ff"><img border="0" src="/image/favor.gif">收藏</font></a></font></p>                                                                                                                                          
                        </td>                                                                                                                                          
                      </tr>                                                                                                                                          
                      <td width="16%"><center>                                                                                                                                           
                      <tr>                                                                                                                                           
                        <td colspan="2" height="9"><img src="/indeximage/line.gif" width="573" height="10"></td>                                                                                                                                          
                      </tr>                                                                                                                                          
                    </table>                                                                                                                                          
                    </td>                                                                                                                                          
                </tr>                                                                                                                                          
              </table>                                                                                                                                          
              <%                                                                                                                                                                                         
	rs.movenext                                                                                                                                                                                         
	loop                                                                                                                                                                                         
	Conn.close                                                                                                                                                                                         
%>                                                                                                                                          
            </td>                                                                                                                                          
          </tr>                                                                                                                                          
        </table>                                                                                                                                                                                
    </td>                                                                                                                                                                                
  </tr>                                                                                                                                                                                
</table>                                                                                                                                                                                
  
</div>                                                                                                                                                                                
  
<map name="Map">                                                                                                                                                                                 
  <area shape="rect" coords="88, 1, 145, 26" href="listen.htm">                                                                                                                                                                                
  <area shape="rect" coords="225, 2, 278, 25" href="xinpin.htm" >                                                                                                                                                                                
  <area shape="rect" coords="157, 2, 210, 25" href="play.htm">                                                                                                                                                                                
</map></body>                                                                                                                                                                                
</html>                                                             