<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
   	dim totalPut   
   	dim CurrentPage
   	dim TotalPages
   	dim i,j
   	dim keyword
   	dim sql
   	dim rs
	dim updown
	dim order_name
	order_name=Request("Order")
   	if not isempty(request("page")) then
      		currentPage=cint(request("page"))
   	else
      		currentPage=1
   	end if
	if request("updown")<>"" then
		updown="desc"
	else
		updown=""
	end if
	select case order_name
		case "showname"
		order_name="showname"
		case "hot"
		order_name="hot"
		case "dateandtime"
		order_name="dateandtime"
		case "hits"
		order_name="hits"
		case "orders"
		order_name="orders"
		case "size"
		order_name="size"
		case else
		order_name="dateandtime"
		updown="desc"
	end select
 	set rs=server.createobject("adodb.recordset")
	dim classid,Nclassid
	dim classname,Nclassname

	if request("classid")="" then
		classid="classid=1 and  "
		sql="select class from class where classid=1"
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
			response.write "��û���κ���Ŀ���뵽����ҳ�����"
			response.end
		else
		classname=rs("class")
		end if
		rs.close
	else
		classid="classid="&cstr(request("classid"))&" and  "
		sql="select class from class where classid="&cstr(request("classid"))
		rs.open sql,conn,1,1
		classname=rs("class")
		rs.close
	end if
	if request("Nclassid")="" then
		Nclassid=""
		Nclassname="����ӰƬ"
	else
		Nclassid=" Nclassid="&cstr(request("Nclassid"))&" and  "
		sql="select Nclass.Nclass,class.class from Nclass,class where Nclass.classid=class.classid and Nclass.Nclassid="&cstr(request("Nclassid"))
		rs.open sql,conn,1,1
		classname=rs("class")
		Nclassname=rs("Nclass")
		rs.close
	end if
%>
<HTML><HEAD><TITLE><%=classname%><%if Nclassid<>"" then%>��<%=Nclassname%><%end if%></TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link rel="stylesheet" href="images/style.css">

<SCRIPT language=javascript>
function popwin2(id,path)
{		window.open("openarticle.asp?id="+id+"&ppath="+path,"");
}
</SCRIPT>
</HEAD>
<!--#include file="topMain.asp"-->
<table width="752" border="0" cellpadding="0" cellspacing="1" align="center" bgcolor="#004C90">
  <tr> 
    <td valign="top" bgcolor="#E8F0F8"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#004C90" height="40">
        <form method="post" name="myform" action="search.asp">
          <tr> 
            <td bgcolor="#E8F0F8" valign="middle" align="center"> ������ 
              <input type="text" name="keyword" class=textfield size=10  maxlength="50"  style="font-family: Arial">
              <input type="submit" name="Submit22" value="����" style="height='21'" class="botton">
            </td>
          </tr>
        </form>
      </table>
      <table border=0 cellpadding=0 width=100% align="center" bgcolor="#004C90" cellspacing=0>
        <tr> 
          <td height="20" bgcolor="#004C90"> 
            <div align="center"><a class=white href='sort.asp?classid=<%=request("classid")%>'><font color="#FFFFFF"><%=classname%></font></a></div>
          </td>
        </tr>
        <tr align=middle bgcolor="#999999"> 
          <td bgcolor="#999999" valign="top"> 
            <table border=0 cellpadding=0 cellspacing=0 width=100% height="100" align="center">
              <tr> 
                <td bgcolor=#E8F0F8         
                  valign=top width="191"><font style=line-height:150%> 
                  <%        
	if request("classid")="" then        
	sql="select Nclass,Nclassid from Nclass where classid=1"        
	else        
	sql="select Nclass,Nclassid from Nclass where classid="&request("classid")        
	end if        
	rs.open sql,conn,1,1%>
                  </font> 
                  <table width="100%" border="0" cellspacing="0">
                    <%do while not rs.eof%>
                    <tr> 
                      <td align="right" height="21"><img src="images/into.gif"></td>
                      <td align="left" height="21"> 
                        <%if not Rs.eof then%>
                        <a href="sort.asp?classid=<%=request("classid")%>&Nclassid=<%=rs("Nclassid")%>"><%=Rs("Nclass")%></a> 
                      </td>
                      <%rs.movenext  
		end if %>
                    </tr>
                    <%loop        
  Rs.Close        
   %>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="20" bgcolor="#004C90"> 
            <p align="center"><font color=#ffffff>������������</font> 
          </td>
        </tr>
        <tr> 
          <td width="100%" bgcolor="#999999"> 
            <table border="0" width="100%" cellspacing="0" cellpadding="0">
              <tr> 
                <td height="20" bgcolor="#E8F0F8" width="4%"></td>
                <td width="96%" bgcolor="#E8F0F8" height="100" valign="top"> 
                  <%if request("Nclassid")="" then
	sql="select top 10 id,showname,bb from download where download.classid="&request("classid")                   
	sql=sql&" order by download.id desc"
	else
	sql="select top 10 id,showname,bb from download where download.Nclassid="&request("Nclassid")              
	sql=sql&" order by download.id desc" 
	end if                  
	rs.open sql,conn,1,1                   
	if rs.eof and rs.bof then
response.write ("û��ӰƬ")
else
do while not rs.eof
response.write ("<li><A href=list.asp?id="&rs("id")&">"&rs("showname")&" "&rs("bb")&"</A></li>")
	rs.movenext                                     
	loop                                     
	end if                                     
	rs.close                                     
%>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td width="180" height="100%" bgcolor="#E8F0F8"></td>
        </tr>
      </table>
    </td>
    <td width="76%" valign="top" bgcolor="#FFFFFF"> 
      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="5"></td>
        </tr>
        <tr>
          <td background="images/bj4.gif" height="1"><img src="images/spacer.gif" width="1" height="1"></td>
        </tr>
      </table>
      <table width="98%" border="0" align="center">
        <tr> 
          <td bgcolor="#efefef">����λ�ã�<a href="index.asp">��ҳ</a> >> <a href="sort.asp?classid=<%=request("classid")%>"><%=classname%></a> 
            >> <%=Nclassname%></td>
        </tr>
      </table>
      
      <table border=0 cellpadding=0 cellspacing=2 width=100%>
        <tbody> 
        <tr> 
          <td valign=top width=582> 
            <%                           
	if request("Nclassid")="" then                          
		sql="select id,showname,bb,note,dayhits,weekhits,lasthits,hits,dateandtime,hot,system,size,orders from download where download.classid="&request("classid")                           
		sql=sql&" order by "&order_name&" "&updown                          
	else                          
		sql="select id,showname,bb,note,dayhits,weekhits,lasthits,hits,dateandtime,hot,system,size,orders from download where download.Nclassid="&request("Nclassid")                          
		sql=sql&" order by "&order_name&" "&updown                          
	end if                          
	rs.open sql,conn,1,1                           
                          
  	if rs.eof and rs.bof then                           
       		response.write "<table border=0 width=100% height=100 cellpadding=0><tr><td width='100%'><p align=center>û�л�û���ҵ��κ���Ϣ</td></tr></table>"                           
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
            		showContent                           
            		showpage totalput,MaxPerPage,"sort.asp"                           
       		else                           
          		if (currentPage-1)*MaxPerPage<totalPut then                           
            			rs.move  (currentPage-1)*MaxPerPage                           
            			dim bookmark                           
            			bookmark=rs.bookmark                           
            			showContent                           
             			showpage totalput,MaxPerPage,"sort.asp"                           
        		else                           
	        		currentPage=1                           
           			showContent                           
           			showpage totalput,MaxPerPage,"sort.asp"                           
	      		end if                           
	   	end if                           
   	rs.close                           
   	end if                           
	                                   
   	sub showContent                           
       	dim i                           
	   	i=0                           
%>
            <table width="98%" >
              <tr>
                <td  width="100%" > 
                  <table border="0" cellspacing="1" cellpadding="0" width="98%" align="center" bgcolor="#3399ff">
                    <%if request("updown")="" then%>
                    <tr> 
                      <td width="37%" align="right" height="22" bgcolor="#E8F0F8"> 
                        <div align="center"><font color="#000000"><b>ӰƬ���ƣ�</b></font></div>
                      </td>
                      <td width="11%" align="right" height="22" bgcolor="#E8F0F8"> 
                        <div align="center"><font color="#000000">ӰƬ����</font></div>
                      </td>
                      <td width="12%" align=center height="22" bgcolor="#E8F0F8"><a href="sort.asp?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=dateandtime&updown=desc" title="�������������"><font color="#000000">��������</font></a></td>
                      <td width="11%" align=center height="22" bgcolor="#E8F0F8"><a href="sort.asp?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=hits&updown=desc"  title="�������������"><font color="#000000">���ش���</font></a></td>
                      <td width="11%" align=center height="22" bgcolor="#E8F0F8"><a href="sort.asp?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=size&updown=desc"  title="�������������"><font color="#000000">�ļ���С</font></a></td>
                      <td width="18%" align=center height="22" bgcolor="#E8F0F8"><a href="sort.asp?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=hot&updown=desc"  title="�������������"><font color="#000000">����</font></a></td>
                    </tr>
                    <%else%>
                    <tr> 
                      <td width="37%" align="right" height="22" background="images/bg.gif"> 
                        <div align="center"><font color="#000000"><b>ӰƬ���ƣ�</b></font></div>
                      </td>
                      <td width="11%" align="right" height="22" bgcolor="#E8F0F8"> 
                        <div align="center"><font color="#000000">ӰƬ����</font></div>
                      </td>
                      <td width="12%" align=center height="22" bgcolor="#E8F0F8"><a href="sort.asp?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=dateandtime" title="�������������"><font color="#000000">��������</font></a></td>
                      <td width="11%" align=center height="22" bgcolor="#E8F0F8"><a href="sort.asp?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=hits"  title="�������������"><font color="#000000">���ش���</font></a></td>
                      <td width="11%" align=center height="22" bgcolor="#E8F0F8"><a href="sort.asp?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=size"  title="�������������"><font color="#000000">�ļ���С</font></a></td>
                      <td width="18%" align=center height="22" bgcolor="#E8F0F8"><a href="sort.asp?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=hot"  title="�������������"><font color="#000000">����</font></a></td>
                    </tr>
                    <%end if%>
                    <%do while not rs.eof%>
                    <tr bgcolor="#FFFFFF" onMouseOver="this.style.backgroundColor='#f8f8f8';" onMouseOut="this.style.backgroundColor='#ffffff';"> 
                      <td width="37%" height="22" bgcolor="#f8f8f8"> <a class="date" href="list.asp?id=<%=rs("id")%>" title="����鿴��ϸ����"><%=rs("showname")%>&nbsp;<%=rs("bb")%></a></td>
                      <td height="22" width="11%"> 
                        <div align="center"><%=rs("orders")%></div>
                      </td>
                      <td width="12%" align=center height="22"><%=month(rs("dateandtime"))%>��<%=day(rs("dateandtime"))%>��</td>
                      <td width="11%" align=center height="22"><%=rs("hits")%></td>
                      <td width="11%" align=center height="22"><%=rs("size")%></td>
                      <td width="18%" align=center height="22"> 
                        <%for imghot=1 to rs("hot")%><img src="images/grade1.gif"><%next%>
                      </td>
                    </tr>
                    <%     
	 i=i+1     
	 if i>=MaxPerPage then exit do     
	 	rs.movenext     
	 loop     
%>
                    <tr bgcolor="#FFFFFF"> 
                      <td colspan="6">&nbsp; </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td colspan="6"> 
                        <p align="center"> 
                          <% end sub      
     
	function showpage(totalnumber,maxperpage,filename)     
  	dim n     
     
  	if totalnumber mod maxperpage=0 then     
     		n= totalnumber \ maxperpage     
  	else     
     		n= totalnumber \ maxperpage+1     
  	end if     
  	response.write "<form method=Post action="&filename&"?classid="&request("classid")&"&Nclassid="&request("Nclassid")&"&order="&request("order")&"&updown="&request("updown")&">"     
  	response.write "<font color='red'>"&Nclassname&"</font>"     
     
  	if CurrentPage<2 then     
    		response.write ""&totalnumber&"��&nbsp;��ҳ ��һҳ&nbsp;"     
  	else     
    		response.write ""&totalnumber&"��&nbsp;<a href="&filename&"?page=1&classid="&request("classid")&"&Nclassid="&request("Nclassid")&"&order="&request("order")&"&updown="&request("updown")&">��ҳ</a>&nbsp;"     
    		response.write "<a href="&filename&"?page="&CurrentPage-1&"&classid="&request("classid")&"&Nclassid="&request("Nclassid")&"&order="&request("order")&"&updown="&request("updown")&">��һҳ</a>&nbsp;"     
  	end if     
     
  	if n-currentpage<1 then     
    		response.write "��һҳ βҳ"     
  	else     
    		response.write "<a href="&filename&"?page="&(CurrentPage+1)&"&classid="&request("classid")&"&Nclassid="&request("Nclassid")&"&order="&request("order")&"&updown="&request("updown")&">"     
    		response.write "��һҳ</a> <a href="&filename&"?page="&n&"&classid="&request("classid")&"&Nclassid="&request("Nclassid")&"&order="&request("order")&"&updown="&request("updown")&">βҳ</a>"     
  	end if     
   	response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&CurrentPage&"</font>/"&n&"</strong>ҳ "     
    	response.write "&nbsp;<b>"&maxperpage&"</b>��ӰƬ/ҳ "     
%>
<%            
end function       
%>
                        </p>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        </tbody> 
      </table><br>
     </td>
  </tr>
</table>
<!--#include file="CopyRight.asp"--> 
</BODY></HTML>