
<%
	FF = request("Flag")
	Num = request("Num")
	if Num = "" then
		SQL = "Select * from 订单管理 order by 订单号"
		set cn=Server.CreateObject("ADODB.Connection")

		cn.open "DBQ="+server.mappath("../btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 

		set rs = cn.Execute(SQL)
	else
		SQL = "Update 订单管理 set 标记='" + FF + "' where 订单号='" + Num + "'"
		set cn=Server.CreateObject("ADODB.Connection")

		cn.open "DBQ="+server.mappath("../btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
		set rs = cn.Execute(SQL)
		SQL = "Select * from 订单管理 order by 订单号"
		set rs = cn.Execute(SQL)
	end if
		

%>


<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>洋洋购物网上购物订单管理系统</title>
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
</head>

<body topmargin="0" leftmargin="0" onLoad="MM_preloadImages('indeximage/gouwu02.gif','indeximage/shoucang02.gif','indeximage/shouyin02.gif','indeximage/kefu02.gif')">
<div align="center">
  <center>
    <table width="768" border="0" cellpadding="0" cellspacing="0">
      <tr> 
        <td background="/indeximage/ditu.gif"> 
          <table width="744" border="0" cellpadding="0" cellspacing="0">
            <tr> 
              <td rowspan="2" width="157"><img src="/indeximage/logo.gif" width="195" height="98"></td>
              <td height="62" colspan="5"><img border="0" src="/indeximage/banner.gif" width="611" height="62"></td>
            </tr>
            <tr> 
              <td width="299"><img src="/indeximage/label1.gif" usemap="#Map" border="0" width="299" height="36"></td>
              <td width="76"> <a href style="CURSOR: hand" onClick="window.open('/shopbag.asp',null,'width=560,height=360,scrollbars=1')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image40','','/indeximage/gouwu02.gif',1)" target="_blank"><img name="Image40" border="0" src="/indeximage/gouwu01.gif" width="76" height="36"></a></td>
              <td width="76"><a href style="CURSOR: hand" onClick="window.open('/favorite.asp',null,'width=560,height=360,scrollbars=1')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image37','','/indeximage/shoucang02.gif',1)" target="_blank"><img name="Image37" border="0" src="/indeximage/shoucang01.gif" width="76" height="36"></a></td>
              <td width="76"><a href style="CURSOR: hand" onClick="window.open('/money.asp',null,'width=560,height=360,scrollbars=1')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image38','','/indeximage/shouyin02.gif',1)" target="_blank"><img name="Image38" border="0" src="/indeximage/shouyin01.gif" width="76" height="36"></a></td>
              <td width="84"><a href="/server.htm" style="CURSOR: hand" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image39','','/indeximage/kefu02.gif',1)"><img name="Image39" border="0" src="/indeximage/kefu01.gif" width="84" height="36"></a></td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
    <map name="Map"> 
      <area shape="rect" coords="90,1,147,26" href="../listen.htm">
      <area shape="rect" coords="225, 2, 278, 25" href="/xinpin.htm" >
      <area shape="rect" coords="157, 2, 210, 25" href="/play.htm">
    </map>
    <table border="0" cellpadding="0" cellspacing="0" width="74%">
    <tr>
      <td width="100%" colspan="5">
        <hr>
      </td>
    </tr>
    <tr>
      <td width="20%" align="center"><font face="方正姚体" size="3" color="#0000FF"><a href="newform.asp">未处理订单</a></font></td>
      <td width="20%" align="center"><font face="方正姚体" size="3" color="#0000FF"><a href="oldform.asp">已处理订单</a></font></td>
      <td width="20%" align="center"><font face="方正姚体" size="3" color="#0000FF"><a href="wasteform.asp">无效订单</a></font></td>
      <td width="20%" align="center"><font face="方正姚体" size="3" color="#0000FF"><a href="formsearch.asp">订单查询</a></font></td>
      <td width="20%" align="center"><font face="方正姚体" size="3" color="#0000FF"><a href="formmanage.asp">订单管理</a></font></td>
    </tr>
    <tr>
      <td width="100%" colspan="5">
        <hr><form name="Modi" action="formmanage.asp" method="POST"><input name="Flag" type="hidden"><input name="Num" type="hidden">
        <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#808080" height="39">
          <tr>
            <td width="3%" align="center" bgcolor="#FFFF66" height="19">&nbsp;</td>
            <td width="12%" align="center" bgcolor="#FFFF66" height="19"><font face="宋体" size="2"><b>订单号码</b></font></td>
            <td width="18%" align="center" bgcolor="#FFFF66" height="19"><font face="宋体" size="2"><b>用户名称</b></font></td>
            <td width="11%" align="center" bgcolor="#FFFF66" height="19"><font face="宋体" size="2"><b>配送方式</b></font></td>
            <td width="11%" align="center" bgcolor="#FFFF66" height="19"><font face="宋体" size="2"><b>付款方式</b></font></td>
            <td width="19%" align="center" bgcolor="#FFFF66" height="19"><font face="宋体" size="2"><b>填表时间</b></font></td>
            <td width="18%" align="center" bgcolor="#FFFF66" height="19"><font face="宋体" size="2"><b>完成时间</b></font></td>
            <td width="8%" align="center" bgcolor="#FFFF66" height="19"><font face="宋体" size="2"><b>标记</b></font></td>
          </tr>
<%
	Do while Not rs.EOF
	if rs.fields("标记") = "废" then
    	FF = "无"
    else
    	FF = "废"
    end if  			
%>
          <tr>
            <td width="3%" align="center" bgcolor="#FFFFFF" height="16"><font size="2" color="blue"><a style="cursor: hand" href onClick="
            	Modi.Flag.value='<%=FF%>';
            	Modi.Num.value='<%=rs.fields("订单号")%>';
            	Modi.submit()
            	">改</a></font></td>
            <td width="12%" align="center" bgcolor="#FFCCFF" height="16"><font face="宋体" size="2"><%=rs.fields("订单号")%></font></td>
            <td width="18%" align="center" bgcolor="#CCFFFF" height="16"><font face="宋体" size="2"><%=rs.fields("用户名")%></font></td>
            <td width="11%" align="center" bgcolor="#FFCCFF" height="16"><font face="宋体" size="2"><%=rs.fields("配送")%></font></td>
            <td width="11%" align="center" bgcolor="#CCFFFF" height="16"><font face="宋体" size="2"><%=rs.fields("付款")%></font></td>
            <td width="19%" align="center" bgcolor="#FFCCFF" height="16"><font face="宋体" size="2"><%=rs.fields("填写时间")%></font></td>
            <td width="18%" align="center" bgcolor="#CCFFFF" height="16"><font face="宋体" size="2"><%=rs.fields("完成时间")%></font></td>
            <td width="8%" align="center" bgcolor="#FFCCFF" height="16"><font face="宋体" size="2"><%=rs.fields("标记")%></font></td>
          </tr>
<%
	rs.Movenext
	loop
	cn.Close
%>
        </table></form>
        <p align="center">Copyright:北京公司</td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>

