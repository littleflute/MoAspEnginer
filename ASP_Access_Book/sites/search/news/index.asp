<!--#include file="articleconn.asp"-->

<html>
<%
   const MaxPerPage=20
   dim totalPut   
   dim CurrentPage
   dim TotalPages
   dim i,j

   dim typename
   dim keyword 
   keyword=trim(request("keyword"))
   typename=request.Querystring("typename")
   if not isempty(request("page")) then
      currentPage=cint(request("page"))
   else
      currentPage=1
   end if
   
%>
<head>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ϣ��ѯϵͳ</title>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<script>
<!--
function openwin(id){		window.open("viewarticle.asp?id="+id,"","height=450,width=550,resizable=yes,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no");
	}
//-->
</script>
<link rel="stylesheet" href="./css/style.css">
</head>
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">
<table width="609" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td> 
      <div align="center"><img src="images/TITLE.jpg" width="587" height="174"></div>
    </td>
  </tr>
</table>
<div align="center"></div>
<table width="607" cellpadding="0" cellspacing="0" align="center">
  <tr> 
    <td colspan="3" height="117"> 
      <table border="1" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td height="16" bgcolor="#B5D85E" width="81%"> 
            <div align="center"> 
              <p><font color="#000000"><strong><%=typename%> �� Ϣ �� ��</strong></font> 
              </p>
            </div>
          </td>
        </tr>
        <tr> 
          <td width="81%" align="center" valign="top" height="139"> 
            <div align="center"> 
              <div id=floater 
style="Z-INDEX: 1; LEFT: 801px; VISIBILITY: visible; WIDTH: 140px; POSITION: absolute; TOP: 175px"> 
                <table class=f9 bordercolor=#ffffff cellspacing=0 bordercolordark=#ffffff 
cellpadding=0 width=127 bordercolorlight=#000000 border=1 valign="center">
                  <tbody> 
                  <tr bgcolor="#B5D85E"> 
                    <td colspan="2"> 
                      <div align="center"><font class="f9" color="#000000">�� Ϣ 
                        �� ��</font><font class=f9 
      color=#ffffff> </font></div>
                    </td>
                  </tr>
                  <tr bgcolor="#FFB74A"> 
                    <td bgcolor="#FFFFFF"> 
                      <div align="center"><font color="#003399">��</font></div>
                    <td bgcolor="#FFFFFF"><a href="index.asp?typename=������Ϣ">���� 
                      ��Ϣ </a> 
                  <tr bgcolor="#FFB74A"> 
                    <td bgcolor="#FFFFFF"> 
                      <div align="center"><font color="#003399">��</font></div>
                    <td bgcolor="#FFFFFF"><a href="index.asp?typename=������Ϣ">���� 
                      ��Ϣ</a> 
                  <tr bgcolor="#FFB74A"> 
                    <td bgcolor="#FFFFFF"> 
                      <div align="center"><font color="#003399">��</font></div>
                    <td bgcolor="#FFFFFF"><a href="index.asp?typename=�ȵ���Ϣ">�ȵ� 
                      ��Ϣ</a> 
                  <tr bgcolor="#B5D85E"> 
                    <td colspan="2"> 
                      <div align="center"><font color="#000000">�� �� �� �� �� Ϣ</font></div>
                    </td>
                  </tr>
                  <tr bordercolordark="#FFFFFF"> 
                    <td height="15"> 
                      <div align="center"><font color="003399">��</font></div>
                    </td>
                    <td height="15"><a href="index.asp?typename=������Ϣ">������Ϣ</a></td>
                  </tr>
                  <tr bordercolordark="#FFFFFF"> 
                    <td> 
                      <div align="center"><font color="003399">��</font></div>
                    </td>
                    <td><a href="index.asp?typename=���������Ϣ">���������Ϣ</a></td>
                  </tr>
                  <tr bgcolor="#B5D85E" bordercolordark="#FFFFFF"> 
                    <td colspan="2"> 
                      <div align="center"><font color="#000000">�� վ �� Ϣ</font></div>
                    </td>
                  </tr>
                  <tr bgcolor="#FFFFFF" bordercolordark="#FFFFFF"> 
                    <td> 
                      <div align="center"><font color="003399">��</font></div>
                    </td>
                    <td><a href="index.asp?typename=վ����Ϣ">վ����Ϣ</a></td>
                  </tr>
                  <tr bgcolor="#FFFFFF" bordercolordark="#FFFFFF"> 
                    <td height="12"> 
                      <div align="center"><font color="003399">��</font></div>
                    </td>
                    <td height="12"><a href="index.asp?typename=վ����Ϣ">վ����Ϣ</a></td>
                  </tr>
                  <tr bgcolor="#B5D85E" bordercolordark="#FFFFFF"> 
                    <td colspan="2"> 
                      <div align="center"><a href="login.asp">�� �� ҳ ��</a> </div>
                    </td>
                  </tr>
                  <tr bgcolor="#FFFFFF"> 
                    <td colspan="2"> 
                      <div align="center"> 
                        <form method="post" action="index.asp?typename=<%=typename%>">
                          <font color="#FF0000"> </font> 
                          <table border=0 cellpadding=0 cellspacing=0 width="70%">
                            <tr> 
                              <td valign=center colspan="3" align="right"> 
                                <div align="center">��ѯ�ؼ���<br>
                                  <font color="#FF0000"><%=keyword%></font> <br>
                                </div>
                              </td>
                            </tr>
                            <tbody> 
                            <tr> 
                              <td width="32%" align="right" rowspan="2"> </td>
                              <td valign=center width="49%" align="center" height="24"> 
                                <input class=TextBorder maxlength=25 
            name=keyword size=15>
                              </td>
                              <td valign=buttom width="19%" align="left" rowspan="2"> 
                                <br>
                              </td>
                            </tr>
                            <tr> 
                              <td valign=center width="49%" align="center"> 
                                <div align="center"> 
                                  <input alt=վ������ border=0 
            name=submit src="ssearch.gif" type=image>
                                  <input alt=reset border=0 
            name=reset src="a.gif" type=image>
                                </div>
                              </td>
                            </tr>
                            </tbody> 
                          </table>
                        </form>
                      </div>
                    </td>
                  </tr>
                  </tbody> 
                </table>
              </div>

              <br>
              <div align="center"> 
                <div align="center"><br>
                  <br>
                  <div align="center"> 
                    <p align="left">
                      <%
if typename="" then
sql="select * from learning where title like '%"&keyword&"%' order by articleid desc"
else
sql="select * from learning where type='"&typename&"' and title like '%"&keyword&"%' order by articleid desc"
end if
dim sql,rs
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
    if rs.eof and rs.bof then
       response.write "<p align='center'> �� û �� �� �� �� Ϣ</p>"
   else
	   totalPut=rs.recordcount
      if currentpage<1 then
          currentpage=1
      end if
      if (currentpage-1)*MaxPerPage>totalput then
	   if (totalPut mod MaxPerPage)=0 then
	     currentpage= totalPut \ MaxPerPage
	   else
	      currentpage= totalPut \ MaxPerPage + 1
	   end if

      end if
       if currentPage=1 then
           showpage totalput,MaxPerPage,"index.asp"
            showContent
            showpage totalput,MaxPerPage,"index.asp"
       else
          if (currentPage-1)*MaxPerPage<totalPut then
            rs.move  (currentPage-1)*MaxPerPage
            dim bookmark
            bookmark=rs.bookmark
           showpage totalput,MaxPerPage,"index.asp"
            showContent
             showpage totalput,MaxPerPage,"index.asp"
        else
	        currentPage=1
           showpage totalput,MaxPerPage,"index.asp"
           showContent
           showpage totalput,MaxPerPage,"index.asp"
	      end if
	   end if
   rs.close
   end if
	        
   set rs=nothing  
   conn.close
   set conn=nothing
  

   sub showContent
       dim i
	   i=0
  
%>
                    </p>
                  </div>
                  <div align="center"> 
                    <center>
                      <table border="1" cellspacing="0" width="64%"
    bgcolor="#F0F8FF" bordercolorlight="#000000" bordercolordark="#FFFFFF" background="../images/top-linebg.gif">
                        <tr> 
                          <td width="13%" align="center" height="20" bgcolor="#B5D85E"><font color="#000000">ID 
                            �� </font></td>
                          <td width="55%" align="center" bgcolor="#B5D85E"><font color="#000000">�� 
                            Ϣ �� ��</font></td>
                          <td width="18%" align="center" bgcolor="#B5D85E"><font color="#000000">�� 
                            �� ʱ ��</font></td>
                          <td width="14%" align="center" bgcolor="#B5D85E"><font color="#000000">�� 
                            ��</font></td>
                        </tr>
                        <%do while not rs.eof%>
                        <tr> 
                          <td width="13%" height="23" bgcolor="#FFFFFF"> 
                            <p align="center"><%=rs("articleid")%> 
                          </td>
                          <td width="55%" bgcolor="#FFFFFF"> 
                            <% if len(rs("title"))>45 then %>
                            <a href="javascript:openwin(<%=rs("articleid")%>)" title="<%=rs("title")%>"><%=left(rs("title"),45)%>...</a> 
                            <% else %>
                            <a href="javascript:openwin(<%=rs("articleid")%>)" title="<%=rs("title")%>"><%=rs("title")%></a> 
                            <% end if %>
                          </td>
                          <td width="18%" bgcolor="#FFFFFF"> 
                            <div align="center"><%=rs("dateandtime")%></div>
                          </td>
                          <td width="14%" bgcolor="#FFFFFF"> 
                            <p align="center"><%=rs("hits")%> 
                          </td>
                        </tr>
                        <% i=i+1
	      if i>=MaxPerPage then exit do
	      rs.movenext
	   loop
		  %>
                      </table>
                    </center>
                  </div>
                  <%
   end sub 

function showpage(totalnumber,maxperpage,filename)
  dim n
  if totalnumber mod maxperpage=0 then
     n= totalnumber \ maxperpage
  else
     n= totalnumber \ maxperpage+1
  end if
  response.write "<form method=Post action='index.asp?typename="&typename&"&keyword="&keyword&"'>"
  response.write "<p align='center' vAlign='bottom'>&gt;&gt;��ҳ&nbsp;"
  if CurrentPage<2 then
    response.write "<font color='999966'>��ҳ ��һҳ</font>&nbsp;"
  else
    response.write "<a href="&filename&"?page=1&typename="&typename&"&keyword="&keyword&">��ҳ</a>&nbsp;"
    response.write "<a href="&filename&"?page="&CurrentPage-1&"&typename="&typename&"&keyword="&keyword&">��һҳ</a>&nbsp;"
  end if
  if n-currentpage<1 then
    response.write "<font color='999966'>��һҳ βҳ</font>"
  else
    response.write "<a href="&filename&"?page="&CurrentPage+1&"&typename="&typename&"&keyword="&keyword&">��һҳ</a>&nbsp;"
    response.write "<a href="&filename&"?page="&n&"&typename="&typename&"&keyword="&keyword&">βҳ</a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&CurrentPage&"</font>/"&n&"</strong>ҳ "
    response.write "&nbsp;��<b>"&totalnumber&"</b>����Ϣ <b>"&maxperpage&"</b>����Ϣ/ҳ "
   response.write " &nbsp;ת����<input class=TextBorder style='TEXT-ALIGN: center' type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"
   response.write " &nbsp;&nbsp;<input alt=ҳ����ת name='submit' src='images/goto.gif' type='image'></span></p></form>"
       
end function

 
%>
                </div>
              </div>
            </div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p align="center">Copyright &copy;���м���������������޹�˾ 2004</p>
</body>
</html>


