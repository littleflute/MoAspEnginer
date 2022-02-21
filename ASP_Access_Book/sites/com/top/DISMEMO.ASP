<HTML><HEAD>
<TITLE>汽配公司</TITLE>

<STYLE type=text/css></STYLE>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script language="javaScript" type="text/javascript"    
      SRC="../top/js/pz_chromeless_2.1.js"></SCRIPT>
     <script language="javaScript" type="text/javascript"
     SRC="../top/js/openScript.js"></script>
	<link rel="stylesheet" type="text/css" href="../css.css">
</HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<TABLE align=center border=0>
  <TBODY>
    <TR> 
      <TD background=../images/001.jpg height=113 width=778>&nbsp; </TD>
    </TR>
	
  </TBODY>
</TABLE>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=778>
  <TBODY>
    <TR> 
      <TD height=39><IMG border=0 height=39 src="../images/002.gif" 
      useMap=#Map2 width=778></TD>
    </TR>
	  <TR> 
      <TD><img src="../images/top.gif" width="776" height="133"> </TD>
    </TR>
  </TBODY>
</TABLE>
               <%
id=request.querystring("id")
%><!-- #include file="connect.asp" --><%
rs.Open "Select * From data where id="&id, conn,3,3%>
<body bgcolor="#FFFFFF" background="../images/bg.gif">
<div align="center"> 
  <table width="776" border="0" cellspacing="0" cellpadding="0" align="center" bordercolordark="#FFFFFF" bordercolorlight="#000000">
    <tr> 
      <td height="16" width="470"><div align="center"><font color="#000000">标题：<%=rs("name")%>[<%=rs("time")%>]</font></div></td>
    </tr>
    <tr> 
      <td > <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr> 
            <td>
              <p><font color="#000000"><br>
                <br>
                <%= Replace(rs("neirong"), vbcrlf,"<br>&nbsp;&nbsp;&nbsp;")%></font></p>
            </td>
          </tr>
        </table><br>
<table width="778" border="0">
          <tr> 
            <td width="778" height="118" background="../images/006.jpg"><div align="center"><font color="#FFFFFF">Copyright&copy;&reg; 
                2003-2004.汽配公司, All Rights Reserved<br>
            </font></div></td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
   <MAP name=Map2>
    <AREA coords=30,9,58,26 
  href="index.asp" shape=RECT>
    <AREA 
  shape=RECT 
  coords=81,8,132,29 href="new.asp">
    <AREA shape=RECT coords=155,10,209,26 
  href="chanpin.asp">
    <AREA 
  shape=RECT 
  coords=231,9,281,26 href="biao.asp">
    <AREA shape=RECT coords=295,9,346,27 
  href="order.asp">
    <AREA 
  shape=RECT 
  coords=371,9,421,27 href="guestbook.asp">
    <AREA 
shape=RECT coords=445,8,504,28 
  href="lianxi.asp">
  </MAP>
</body>
</html>