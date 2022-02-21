<!--#include file="conn.asp"-->
<!--#include file="home1.asp"-->
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
'		classid=""
'		classname="所有软件"
		classid="classid=1 and  "
		sql="select class from class where classid=1"
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
			response.write "还没有任何栏目，请到管理页面添加"
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
		Nclassname="所有软件"
	else
		Nclassid=" Nclassid="&cstr(request("Nclassid"))&" and  "
		sql="select Nclass.Nclass,class.class from Nclass,class where Nclass.classid=class.classid and Nclass.Nclassid="&cstr(request("Nclassid"))
		rs.open sql,conn,1,1
		classname=rs("class")
		Nclassname=rs("Nclass")
		rs.close
	end if

%> <HTML>
<HEAD>
<TITLE><%=classname%> 
<%if Nclassid<>"" then%>
：<%=Nclassname%> 
<%end if%>
</TITLE>
<META content="text/html; charset=gb2312" http-equiv="Content-Type">
</HEAD>
<!--#include file="head.asp"-->
<table width="770" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#6AAC91"><div align="center">
        <!--#include file="lanmu.asp"-->
    </div></td>
  </tr>
</table>
<table width="770"  border=0 cellspacing="0" cellpadding="0">
  <tr> 
    <td height="22" bgcolor="#FFFFFF" class="4b">当前位置：<a href="<%=rs1("url")%>"><%=rs1("home")%></a>>> 
      <a href="<%=rs1("urls")%>"><%=rs1("homes")%></a>>> <a href="class.asp?classid=<%=request("classid")%>"><%=classname%></a> 
    >> <%=Nclassname%> </td>
  </tr>
</table>
<br>
<table border="0" cellpadding="1" cellspacing="0"  width="770" id="AutoNumber1">
  <tr> 
    <td width="170" valign="top" bgcolor="#FFFFFF"> <table border=0 cellpadding="0" cellspacing="0"  width="168" height="73">
        <tr> 
          <td width="174" height="22" class="li"> 搜索引擎</td>
        </tr>
        <tr bgcolor=#ffffff> 
          <td width="174" align="center" class="3b" background="pic/a-27.gif" bgcolor="#ffffff"><TABLE width="98%" height="63" border=0 align=center cellPadding=0 cellSpacing=2>
              <FORM action=query.asp method=POST name="myfrom">
                <TR> 
                  <TD> <input onfocus="this.value=''" maxLength="50" size="17" title="输入关键字" value="输入关键字" name="keyword" type="text"> 
                  </TD>
                </TR>
                <TR> 
                  <TD><input type="radio" value="title" checked name="action">
                    名称 
                    <input type="radio" name="action" value="content">
                    简介<font color="#FFFFFF">&nbsp;</font></TD>
                </TR>
                <TR> 
                  <TD><input type=submit value="搜 索" name=Submit> <a href="search.asp"> 
                    高级搜索...</a></TD>
                </TR>
              </FORM>
            </TABLE></td>
      </table>
      <br>
      <table width="168"  border=0 cellpadding="0" cellspacing="0">
        <tr> 
          <td class="li" height="22"><%=classname%>分类</td>
        </tr>
        <tr> 
          <td class="3b" background="pic/a-27.gif"> 
            <%        
	if request("classid")="" then        
	sql="select Nclass,Nclassid from Nclass where classid=1"        
	else        
	sql="select Nclass,Nclassid from Nclass where classid="&request("classid")        
	end if        
	rs.open sql,conn,1,1%>
            <table width="100%" border="0" cellspacing="0"  cellpadding="0">
              <%do while not rs.eof%>
              <tr> 
                <td align="right" width="10%"  >·</td>
                <td width="90%"> 
                  <%if not Rs.eof then%>
                  <span style="letter-spacing: 2pt"> <a href="class.asp?classid=<%=request("classid")%>&Nclassid=<%=rs("Nclassid")%>"><%=Rs("Nclass")%></a></span> 
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
      <br>
      <table width="168"  border=0 cellpadding="0" cellspacing="0">
        <tr> 
          <td class="li" height="22"><%=classname%>排行</td>
        </tr>
        <tr> 
          <td background="pic/a-27.gif" class="3b"> 
            <%if request("Nclassid")="" then
	sql="select top 10 id,showname,bb from download where download.classid="&request("classid")                   
	sql=sql&" order by download.id desc"
	else
	sql="select top 10 id,showname,bb from download where download.Nclassid="&request("Nclassid")              
	sql=sql&" order by download.id desc" 
	end if       
	rs.open sql,conn,1,1                   
	if rs.eof and rs.bof then
response.write ("没有源码")
else
do while not rs.eof
response.write ("·<A href=software.asp?id="&rs("id")&">"&rs("showname")&"</A><br>")
	rs.movenext                                     
	loop                                     
	end if                                     
	rs.close                                     
%>
          </td>
        </tr>
      </table>
    </td>
    <td width="596" align="right" valign="top" bgcolor="#FFFFFF"> 
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
       		response.write "<table border=0 width=100% height=100 cellpadding=0><tr><td width='100%'><p align=center>没有或没有找到任何程序</td></tr></table>"                           
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
            		showpage totalput,MaxPerPage,"class.asp"                           
       		else                           
          		if (currentPage-1)*MaxPerPage<totalPut then                           
            			rs.move  (currentPage-1)*MaxPerPage                           
            			dim bookmark                           
            			bookmark=rs.bookmark                           
            			showContent                           
             			showpage totalput,MaxPerPage,"class.asp"                           
        		else                           
	        		currentPage=1                           
           			showContent                           
           			showpage totalput,MaxPerPage,"class.asp"                           
	      		end if                           
	   	end if                           
   	rs.close                           
   	end if                           
	                                   
   	sub showContent                           
       	dim i                           
	   	i=0                           
%>
      <%if request("updown")="" then%>
      <table width="595" bordercolordark=#ffffff  bordercolorlight=#18A6FF border=1 cellspacing="0" cellpadding="0" height="22" background="pic/bar.gif">
        <tr> 
          <td height="14" width="258">&nbsp;&nbsp; 软件名称</td>
          <td width="87" align="center" height="14">更新日期</td>
          <td width="62" align="center" height="14">下载次数</td>
          <td width="71" align="center" height="14">文件大小</td>
          <td width="105" align="center" height="14">软件评级</td>
        </tr>
      </table>
      <table bordercolordark=#ffffff bordercolorlight=#18A6FF border=1 width="595" cellspacing="0" cellpadding="0">
        <%end if%>
        <%do while not rs.eof%>
        <tr> 
          <td width="258" height="22"> <img src="images/isList.gif" width="13" height="13"> 
            <a href="software.asp?id=<%=rs("id")%>" title="<%=left(rs("note"),150)%>"><b><font color="#000000"><%=rs("showname")%></font></b> 
            </a></td>
          <td width="87" align="center" height="22">2003-<%=month(rs("dateandtime"))%>-<%=day(rs("dateandtime"))%></td>
          <td width="62" align="center" height="22"><%=rs("hits")%></td>
          <td width="71" align="center" height="22"><%=rs("size")%></td>
          <td width="105" align="center" height="22"> 
            <%for imghot=1 to rs("hot")%>
            ★ 
            <%next%>
          </td>
        </tr>
        <tr> 
          <td colspan="5">&nbsp;&nbsp; · <%=left(rs("note"),150)%>...... </td>
        </tr>
        <tr> 
          <td colspan="3" height="23"><font color="green">运行平台：</font><%=rs("system")%> 
            <font color="green">软件性质：</font><%=rs("orders")%>/<%=Nclassname%></td>
          <td colspan="2" height="23"><font color="green">人气：</font> <%=rs("hits")%>℃ 
          </td>
        </tr>
      </table>
      <br>
      <table width="595" bordercolordark=#ffffff bordercolorlight=#18A6FF border=1 cellspacing="0" cellpadding="0">
        <%     
	 i=i+1     
	 if i>=MaxPerPage then exit do     
	 	rs.movenext     
	 loop     
%>
        <tr> 
          <td> 
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
    		response.write ""&totalnumber&"个&nbsp;首页 上一页&nbsp;"     
  	else     
    		response.write ""&totalnumber&"个&nbsp;<a href="&filename&"?page=1&classid="&request("classid")&"&Nclassid="&request("Nclassid")&"&order="&request("order")&"&updown="&request("updown")&">首页</a>&nbsp;"     
    		response.write "<a href="&filename&"?page="&CurrentPage-1&"&classid="&request("classid")&"&Nclassid="&request("Nclassid")&"&order="&request("order")&"&updown="&request("updown")&">上一页</a>&nbsp;"     
  	end if     
     
  	if n-currentpage<1 then     
    		response.write "下一页 尾页"     
  	else     
    		response.write "<a href="&filename&"?page="&(CurrentPage+1)&"&classid="&request("classid")&"&Nclassid="&request("Nclassid")&"&order="&request("order")&"&updown="&request("updown")&">"     
    		response.write "下一页</a> <a href="&filename&"?page="&n&"&classid="&request("classid")&"&Nclassid="&request("Nclassid")&"&order="&request("order")&"&updown="&request("updown")&">尾页</a>"     
  	end if     
   	response.write "&nbsp;页次：<strong><font color=red>"&CurrentPage&"</font>/"&n&"</strong>页 "     
    	response.write "&nbsp;<b>"&maxperpage&"</b>个软件/页 "     
%>
            <%            
end function       
%>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>
<!--#include file="foot.asp"-->
</BODY>
</HTML>