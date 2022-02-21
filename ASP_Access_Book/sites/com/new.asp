<!-- #include file="conn.asp" -->
<HTML><HEAD>
<TITLE>汽配制造有限公司</TITLE>
<STYLE type=text/css></STYLE>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script language="JavaScript">
function NewsWindow(id)
{
window.open('new/newswind.asp?id='+id,'infoWin', 'height=400,width=600,scrollbars=yes,resizable=yes');
}
</script>
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
              <TD bgColor=#ffe6c2 vAlign=top width=212><img src="images/q.gif" width="200" height="250"> 
              </TD>
              <TD vAlign=top width=31>&nbsp;</TD>
              <TD vAlign=top width=521><table width="100%" border="0">
                  <tr>
                    <td> 
                      <%
set rs=server.createobject("adodb.recordset")
sql="SELECT * from news order by ID desc"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "<p>还 没 有 任 何 新 闻</p>"
else%>
                      <p> 
                        <%
cc=1
	if not isempty(request("page")) then
		pagecount=cint(request("page"))
	else
		pagecount=1
	end if
	rs.PageSize=10
	rs.AbsolutePage=pagecount
For iPage = 1 To rs.PageSize
If rs.EOF Then Exit For
if cc mod 2=1 then
	Response.Write "<tr bgcolor=#E7E7E7>"
else
	Response.Write "<tr BGCOLOR=#F4F4F4>"
end if
%>
                        <img src=images/dot.gif width="15" height="11"><a href="new/newswind.asp?id=<%=rs("ID")%>" target=_blank><u><%=rs("title")%></u></a> 
                        <font size="1"><%=rs("times")%></font> <br>
                        <%
if DateDiff("d",rs("times"),date())<1 then Response.Write ""
Response.Write "</td></tr>"
cc=cc+1
rs.movenext
Next
Response.Write "</table><p>共"&rs.recordcount&"条新闻&nbsp;"
if rs.PageCount>1 Then
 If pagecount<>1 Then
   Response.Write "<A HREF=new.asp?Page=1>首页&nbsp;</A>" 
   Response.Write "<A HREF=new.asp?Page="&(pagecount-1)&">前页&nbsp;</A>"
 End If
 If pagecount<>rs.PageCount Then
   Response.Write "<A HREF=new.asp?Page="&(pagecount+1)&">后页&nbsp;</A>"
   Response.Write "<A HREF=new.asp?Page="&rs.PageCount&">尾页&nbsp;</A>"
 End If
End If
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
                      </p>
</td>
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