<!--#include file="conn.asp"-->
<HTML><HEAD>
<TITLE>汽配制造有限公司</TITLE>
<STYLE type=text/css></STYLE>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>

	<link rel="stylesheet" type="text/css" href="css.css">
</HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<TABLE align=center border=0>
  <TBODY>
    <TR> 
      <TD background=images/001.jpg height=113 width=778>&nbsp; </TD>
    </TR>
  </TBODY>
</TABLE>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=778>
  <TBODY>
    <TR> 
      <TD height=39><IMG border=0 height=39 src="images/002.gif" 
      useMap=#Map2 width=778></TD>
    </TR>
    <TR> 
      <TD><img src="images/top.gif" width="776" height="133"> </TD>
    </TR>
  </TBODY>
</TABLE>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=778>
  <TBODY>
    <TR> 
      <TD bgColor=#f0f0f0 vAlign=top> <TABLE height="172" border=0>
          <TBODY>
            <TR borderColor=#000000> 
              <TD bgColor=#ffe6c2 vAlign=top width=212><img src="images/c.gif" width="200" height="250"> 
              </TD>
              <TD vAlign=top width=31>&nbsp;</TD>
              <TD vAlign=top width=521><table width="100%" border="0">
                  <tr>
                    <td><div align="center">
                        <center>
                          <table border="0" width="500" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td width="100%"> 
                                <% 
      dim pic_class
      pic_class=trim(request("p_name"))
      
      
      set rst=server.createobject("adodb.recordset")
      if pic_class="" then
      sql="select * from pic order by id desc"
      else
      sql="select * from pic where pic_class='"&pic_class&"' order by id desc"
      end if
      rst.open sql,conn,3,1      
	if Not(rst.bof and rst.eof) then'判别数据表中是否为空记录
			NumRecord=rst.recordcount
			rst.pagesize=2
			NumPage=rst.Pagecount
			if request("page")=empty then 
			NoncePage=1
		else
		if Cint(request("page"))<1 then
			NoncePage=1
		else
			NoncePage=request("page")
		end if
		if Cint(Trim(request("page")))>Cint(NumPage) then NoncePage=NumPage
	end if
else
	NumRecord=0
	NumPage=0
	NoncePage=0
	end if
%>
                                <div align="center"> 
                                  <table border="0" width="100%" bordercolorlight="#000000" cellspacing="0" cellpadding="2" bordercolordark="#FFFFFF">
                                    <center>
                                      <%if Not(rst.bof and rst.eof) then
	rst.move (Cint(NoncePage)-1)*2,1
	for i=1 to rst.pagesize
%>
                                      <tr> 
                                        <td> <table border="0" width="100%" cellspacing="0" cellpadding="5" height="232">
                                            <tr> 
                                              <td width="42%" rowspan="7" height="134"> 
                                                <p align="center"><a target="_blank" href="show_pic.asp?id=<%=rst("id")%>"><img src="<%=rst("pic_lurl")%>" alt="点击放大图片" width="194" height="134" border="0"></a></td>
                                              <td width="16%" align="right" height="28">产品名称：</td>
                                              <td width="42%" height="28"><%=rst("pic_name")%></td>
                                            </tr>
                                            <tr>
                                              <td align="right" height="14">产品价格：</td>
                                              <td height="14">500元/件</td>
                                            </tr>
                                            <tr> 
                                              <td width="16%" align="right" height="14">产品类型：</td>
                                              <td width="42%" height="14"><%=rst("pic_class")%></td>
                                            </tr>
                                            <tr> 
                                              <td width="16%" align="right" height="14">产品型号：</td>
                                              <td width="42%" height="14"><%=rst("pic_type")%></td>
                                            </tr>
                                            <tr> 
                                              <td width="16%" align="right" height="14">产品介绍：</td>
                                              <td width="42%" height="14"><%=rst("pic_text")%></td>
                                            </tr>
                                            <tr> 
                                              <td width="16%" align="right" height="14">点击数：</td>
                                              <td width="42%" height="14"><%=rst("pic_count")%></td>
                                            </tr>
                                            <tr> 
                                              <td width="16%" align="right" height="14">发布时间：</td>
                                              <td width="42%" height="14"><%=rst("pic_time")%></td>
                                            </tr>
                                            <tr> 
                                              <td height="14" colspan="3"> <hr>                                              </td>
                                            </tr>
                                          </table></td>
                                      </tr>
                                      <%		 		rst.movenext
			   		if rst.eof then exit for
					next
else
	response.write "<tr><td colspan=13><marquee scrolldelay=120 behavior=alternate>没有找到任何记录!!!</marquee></td></tr>"
end if	

rst.close
set rst=nothing

%>
                                    </center>
                                  </table>
                                </div>
                          </table>
                        </center>
                      </div>
                      <table width="550" border="0" align="center">
                        <tr> 
                          <td height="17"> <div align="right"> 
                              <input type="hidden" name="page" value="<%=NoncePage%>">
                              <%
if NoncePage>1 then
	response.write "|<a href=chanpin.asp?page=1>首 页</a>| |<a href=chanpin.asp?page="&NoncePage-1&">上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=chanpin.asp?page="&NoncePage+1&">下一页</a>| |<a href=chanpin.asp?page="&NumPage&">尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
                              &nbsp;页次：<font color="#0033CC"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 
                              共<font color="#0033CC"><%=NumRecord%></font>条记录&nbsp; 
                            </div></td>
                        </tr>
                      </table>
                      <p></p></td>
                  </tr>
                </table>
                <p>&nbsp;</p></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
  </TBODY>
</TABLE>
<table width="778" border="0" align="center">
  <tr> 
    <td width="778" height="118" background="images/006.jpg"><div align="center"><font color="#FFFFFF">Copyright&copy;&reg; 
        2003-2004.汽配制造有限公司, All Rights Reserved<br>
        </font></div></td>
  </tr>
</table>
<div align="center">
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
</div>
</BODY></HTML>