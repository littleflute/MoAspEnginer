<!--#include file="adovbs.inc" -->
<!--#include file="ShowPage.asp" -->
<!--#include file="Functions.asp" -->
<%
	if request("P") <> "" then
		'T= "where ���� like '" + request("P") + "%'"
	else
		'T= "where ����='" + request("C") + "'"
	end if
%>


<html>
<head>
<title>�����������</title>
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
	COLOR: #282828; FONT-FAMILY: ����; FONT-SIZE: 9pt; LINE-HEIGHT: 16px
}
</STYLE>



</head>

<body bgcolor="#006699" text="#000000" onLoad="MM_preloadImages('indeximage/gouwu02.gif','indeximage/shoucang02.gif','indeximage/shouyin02.gif','indeximage/kefu02.gif')" topmargin="0" leftmargin="2">
<div align="center"><center> 
  <table width="768" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td rowspan="2" width="157"><img src="indeximage/logo.gif" width="157" height="98"></td>
      <td height="62" colspan="6" width="613"><img src="indeximage/banner.gif" width="611" height="62"></td>
    </tr>
    <tr> 
      <td width="299"><img src="indeximage/label1.gif" width="299" height="36" usemap="#Map" border="0"></td>
      <td width="76"> <a href style="CURSOR: hand" onClick="window.open('shopbag.asp',null,'width=560,height=360,scrollbars=1')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image40','','indeximage/gouwu02.gif',1)" target="_blank"><img name="Image40" border="0" src="indeximage/gouwu01.gif" width="76" height="36"></a></td>
      <td width="76"><a href style="CURSOR: hand" onClick="window.open('favorite.asp',null,'width=560,height=360,scrollbars=1')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image37','','indeximage/shoucang02.gif',1)" target="_blank"><img name="Image37" border="0" src="indeximage/shoucang01.gif" width="76" height="36"></a></td>
      <td width="76"><a href style="CURSOR: hand" onClick="window.open('money.asp',null,'width=560,height=360,scrollbars=1')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image38','','indeximage/shouyin02.gif',1)" target="_blank"><img name="Image38" border="0" src="indeximage/shouyin01.gif" width="76" height="36"></a></td>
      <td width="84"><a href="server.htm" style="CURSOR: hand" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image39','','indeximage/kefu02.gif',1)"><img name="Image39" border="0" src="indeximage/kefu01.gif" width="84" height="36"></a></td>
    </tr>
  </table>
  <map name="Map"> 
    <area shape="rect" coords="155, 4, 212, 24" href="play.htm">
    <area shape="rect" coords="90, 4, 143, 24" href="listen.htm">
    <area shape="rect" coords="227, 3, 280, 23" href="/xinpin.htm" >
    <area shape="rect" coords="22, 4, 82, 23" href="index.asp">
  </map>
  <table width="720" border="0" cellspacing="0" cellpadding="0">
    <tr> 
    <td width="169" height="654" valign="top" background="image/ditu07.gif"> 
      <table width="167" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="34%" valign="top" background="image/ditu01.gif"> 
              <form name="Searching" method="POST" action="SearchResult.asp">
          
          
                <table width="99%" border="0" cellspacing="0" cellpadding="2">
                  <tr> 
                    <td width="30%"><font size="2"><b>���:</b></font></td>
                    <td width="70%" colspan="2"> 
                      <select size="1" name="D">
                        <option selected value="��ƪ����">��ƪ����</option>
                        <option value="��ͳ����">��ͳ����</option>
                        <option value="��������">��������</option>
                        <option value="��������">��������</option>
                        <option value="����ѧϰ">����ѧϰ</option>
                        <option value="����ѧϰ">����ѧϰ</option>
                        <option value="��ͥ�ٿ�">��ͥ�ٿ�</option>
                        <option value="��ý����">��ý����</option>
                        <option value="�鶾ɱ��">�鶾ɱ��</option>
                        <option value="ϵͳӦ��">ϵͳӦ��</option>
                        <option value="��ɫ����">��ɫ����</option>
                        <option value="����ð��">����ð��</option>
                        <option value="ս�Բ���">ս�Բ���</option>
                        <option value="ģ�⾭Ӫ">ģ�⾭Ӫ</option>
                        <option value="��������">��������</option>
                        <option value="��������">��������</option>
                        <option value="�������">--�������-</option>
                      </select>
                    </td>
                  </tr>
                  <tr> 
                    <td width="30%" height="22"><font size="2"><b>�ִ�:</b></font></td>
                    <td width="39%" height="22"> 
                      <input type="text" name="StrL" size="7">
                    </td>
                    <td width="31%" height="22" valign="bottom" align="right"><a href style="CURSOR:hand" onClick="
                    	if (Searching.StrL.value == '')
                    		{
                    			alert('��������д��ѯ������!')
                    		}
                    	else
                    		{
                    			Searching.submit()
                    		}">
						<img src="indeximage/search.gif" width="38" height="19"></a></td>
                  </tr>
                </table></form>
            
            
          </td>
          <td width="66%" valign="top"><img src="image/001.gif" width="14" height="78"></td>
        </tr>
        <tr> 
          <td valign="top" width="34%"><img src="image/book01.gif" width="155" height="32"></td>
          <td valign="top" width="66%"><img src="image/002.gif" width="14" height="32"></td>
        </tr>
        <tr> 
            <td width="34%" background="image/ditu02.gif" valign="top" height="85"> 
              <table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="7%" valign="top"><img src="image/ani_arrow1.gif" width="22" height="17"></td>
                  <td width="93%"><font size="2"><b><font color="#336600">��ƪ����</font></b></font></td>
                </tr>
                <tr> 
                  <td width="7%">&nbsp;</td>
                  <td width="93%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?P=ĳĳ'">ĳĳ</a>  
                    ����<a href style="CURSOR: hand" onClick="this.href='book.asp?P=ĳĳ'">ĳĳ</a></font></td>
                </tr>
                <tr> 
                  <td width="7%">&nbsp;</td>
                  <td width="93%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?P=ĳĳ'">ĳĳ</a>  
                    ����<a href style="CURSOR: hand" onClick="this.href='book.asp?P=ĳĳ'">ĳĳ</a></font></td>
                </tr>
                <tr> 
                  <td width="16%">&nbsp;</td>
                  <td width="84%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?P=ĳĳ'">ĳĳ</a></font></td>
                </tr>
                <tr> 
                  <td colspan="2"><font size="2"><b>------&lt;&lt;&lt;&gt;&gt;&gt;------</b></font></td>
                </tr>
              </table>
          </td>
            <td width="66%" valign="top" background="image/004.gif" height="85"><img src="image/004.gif" width="14" height="99"></td>
        </tr>
        <tr> 
          <td width="34%" background="image/ditu04.gif" valign="top" height="49"> 
            <table width="100%" border="0" cellspacing="0">
              <tr> 
                <td width="7%" valign="top"><img src="image/ani_arrow1.gif" width="22" height="17"></td>
                <td width="93%"><font size="2"><b><font color="#336600">��ͳ����</font></b></font></td>
              </tr>
              <tr> 
                <td width="7%">&nbsp;</td>
                <td width="93%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?P=ĳĳ'">ĳĳ</a></font></td>
              </tr>
              <tr> 
                <td width="7%">&nbsp;</td>
                <td width="93%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?P=ĳĳ'">ĳĳ</a></font></td>
              </tr>
              <tr> 
                <td width="7%">&nbsp;</td>
                <td width="93%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?P=ĳĳ'">ĳĳ</a></font></td>
              </tr>
              <tr> 
                <td width="7%">&nbsp;</td>
                <td width="93%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?P=ĳĳ'">ĳĳ</a></font></td>
              </tr>
            </table>
          </td>
          <td width="66%" valign="top" background="image/005.gif" height="49"><img src="image/005.gif" width="14" height="99"></td>
        </tr>
        <tr> 
          <td width="34%" background="image/bookditu02.gif"> 
            <table width="100%" border="0" cellspacing="0">
              <tr> 
                <td width="15%"><img src="image/ani_arrow1.gif" width="22" height="17"></td>
                <td width="85%"><font color="#336600" size="2"><b><a href style="CURSOR: hand" onClick="this.href='book.asp?P=��������'">��������</a></b></font></td>
              </tr>
              <tr> 
                <td width="15%">&nbsp;</td>
                <td width="85%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?C=վ�ڲ��Ա���'">վ�ڲ��Ա���</a></font></td>
              </tr>
              <tr> 
                <td width="15%">&nbsp;</td>
                <td width="85%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="self.location.replace('book.asp?P=վ�ڲ��Ա���')">վ�ڲ��Ա���</a></font></td>
              </tr>
              <tr> 
                <td width="15%">&nbsp;</td>
                <td width="85%"><font size="2" color="0000ff"><a href style="CURSOR: hand" onClick="this.href='book.asp?P=վ�ڲ��Ա���'">վ�ڲ��Ա���</a></font></td>
              </tr>
              <tr> 
                <td colspan="2"><font size="2"><b>------&lt;&lt;&lt;&gt;&gt;&gt;------</b></font></td>
              </tr>
            </table>
          </td>
          <td width="66%" valign="top" background="image/ditu06.gif"><img src="image/game02.gif" width="14" height="100"></td>
        </tr>
        <tr> 
          <td width="34%" background="image/bookditu01.gif" height="29">&nbsp;</td>
          <td width="66%" valign="top" background="image/ditu06.gif" height="29"><img src="image/game02.gif" width="14" height="100"></td>
        </tr>
      </table>
    </td>
      <td width="551" height="654" valign="top" bgcolor="#ffffff"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="2" height="172">
            <tr> 
              
            <td height="18" valign="top"><img src="indeximage/13.gif" width="595" height="9"></td>
            </tr>
            <tr> 
              <td height="7" valign="top"> 
              
          
<%               
	SQL="Select * from ��Ʒ��Ϣ " + T
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
              
<form action="book.asp" method="POST">              
  <div align="right"><p>
<%               
   If Page <> 1 Then              
      Response.Write "<A HREF=book.asp?Page=1><font size=2 color=blue>��һҳ</font></A>"              
      Response.Write "&nbsp; <A HREF=book.asp?Page=" & (Page-1) & "><font size=2 color=blue> ��һҳ</font></A>"              
   End If              
   If Page <> rs.PageCount Then              
      Response.Write "&nbsp;<A HREF=book.asp?Page=" & (Page+1) & "><font size=2 color=blue>��һҳ</font></A>"              
      Response.Write "&nbsp; <A HREF=book.asp?Page=" & rs.PageCount & "><font size=2 color=blue>��ĩҳ</font></A>&nbsp;"              
   End If  
   
%>              
    <font size="2"> ����<input type="text" size="3" name="Page">                                                                        
    ҳ��</font><font color="#FF0000"><%=Page%>/<%=rs.PageCount%></font> </p>                                                                        
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
  </table>                                         
<table width="771">                                                            
  <tr>                                                        
    <td width="763">                                                        
      <hr width="100%" size="1" noshade>                                                       
    </td>                                                       
  </tr>                                                       
  <tr>                                                       
    <td width="763" height="4" align="center">&nbsp;<font size="2"><b>Copyright</b>:֣������</font></td>                                                       
  </tr>                                                       
  
</table>    
</div> 
</body>    
</html>

