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
	dim founderr
	dim errmsg
	dim findword
	founderr=false
	keyword=request("keyword")
	if keyword="" then
		errmsg=errmsg+"<br>"+"�������ѯ������"
		founderr=true
	else
		if request("action")="title" then
			findword="showname like '%"&keyword&"%' "
		else
			findword="note like '%"&keyword&"%' "
		end if
	end if
   	if not isempty(request("page")) then
      		currentPage=cint(request("page"))
   	else
      		currentPage=1
   	end if
 	set rs=server.createobject("adodb.recordset")
	dim classid,Nclassid
	dim classname,Nclassname

	if request("classid")="" then
		classid=""
		classname="����ӰƬ"
	else
		classid="classid="&cstr(request("classid"))&" and  "
		sql="select class from class where classid="&cstr(request("classid"))
		rs.open sql,conn,1,1
		classname="[<font color=#008000>"&rs("class")&"</font>]"
		rs.close
	end if
	if request("Nclassid")="" then
		Nclassid=""
		Nclassname="����ӰƬ"
	else
		Nclassid=" Nclassid="&cstr(request("Nclassid"))&" and  "
		sql="select Nclass from Nclass where Nclassid="&cstr(request("Nclassid"))
		rs.open sql,conn,1,1
		Nclassname="��[<font color=#008000>"&rs("Nclass")&"</font>]����"
		rs.close
	end if

%>
<HTML><HEAD><TITLE>���ҽ��(<%=keyword%>)</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link rel="stylesheet" href="images/style.css">

<SCRIPT language=javascript>
function popwin2(id,path)
{		window.open("openarticle.asp?id="+id+"&ppath="+path,"");
}
</SCRIPT>
</HEAD>
<!--#include file="topMain.asp"-->
<div align="center">
  <center>
  </center>      
</div>      
<%if founderr=true then%>
<table border="0" width="752" height="300" cellpadding="0" cellspacing="1" bgcolor="#004C90" align="center">
  <tr>
    <td width="24%" bgcolor="#E8F0F8" valign="top"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#004C90" height="40">
        <form method="post" name="myform" action="search.asp">
          <tr> 
            <td bgcolor="#E8F0F8" valign="middle" align="center"> ������ 
              <input type="text" name="keyword" class=textfield size=10  maxlength="50"  style="font-family: Arial">
              <input type="submit" name="Submit" value="����" style="height='21'" class="botton">
            </td>
          </tr>
        </form>
      </table>
    </td>
    <td width="76%" bgcolor="#FFFFFF" valign="top"> 
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
          <td>����λ�ã�<a href="index.asp">��ҳ >></a> 
            <%     
	response.write "��ѯ������<font color=red>"&keyword&"</font>"     
	response.write "��"&classname&""&Nclassname&"��"     
%>
          </td>
        </tr>
      </table>
      <br><br><br><br><br><br><br>
      <div align="center">�ؼ��ֲ���Ϊ��!</div>
    </td>
  </tr>
</table>     
<%else%>     
<table border=0 cellpadding=0 cellspacing=1 width=752 align="center">
  <tr>    
    <td valign="top" width="182" bgcolor="#E8F0F8"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#004C90" height="40">
        <form method="post" name="myform" action="search.asp">
          <tr> 
            <td bgcolor="#E8F0F8" valign="middle" align="center"> ������ 
              <input type="text" name="keyword" class=textfield size=10  maxlength="50"  style="font-family: Arial">
              <input type="submit" name="Submit" value="����" style="height='21'" class="botton">
            </td>
          </tr>
        </form>
      </table>
      <TABLE border=0 cellpadding=0 cellspacing=0 width=182 align="center">
        <TR> 
          <TD bgcolor=#004C90 height="20"> 
            <div align="center"><FONT color=#ffffff>��������Top10</FONT></div>
          </TD>
        </TR>
        <TR align=middle> 
          <TD bgcolor=#999999 valign="top"> 
            <TABLE border=0 cellpadding=0 cellspacing=0 width=100% height="100">
              <TR> 
                <TD bgcolor=#E8F0F8 valign=top width="191"><FONT style=line-height:150%> 
                  <%                                       
	dim tdate                                       
	tdate=year(Now()) & "-" & month(Now()) & "-" & day(Now())                                       
	sql="select * from download where "                                       
	sql=sql&" lasthits="&tdate&"  and dayhits>0 "                                       
	sql=sql&" order by dayhits desc"                                       
	rs.open sql,conn,1,1                                       
	if rs.eof and rs.bof then                                       
%>
                  ����û������ 
                  <%else%>
                  <%do while not rs.eof                                         
response.write "<LI><A href=list.asp?id="&rs("id")&">"&rs("showname")&" "&rs("bb")&"</A></LI>"                                         
i=i+1       
if i>=10 then exit do       
rs.movenext     
loop     
i="0"     
	end if 
	rs.close 
%>
                  </FONT> </TD>
              </TR>
            </TABLE>
          </TD>
        </TR>
      </TABLE>                                        
	  
    </td>   
    <td valign=top width=568 height="160" bgcolor="#FFFFFF"> 
      <div align="right"> 
        <table width="98%" border="0" cellpadding="0" cellspacing="0" align="center">
          <tr>
            <td height="5"></td>
          </tr>
          <tr>
            <td background="images/bj4.gif" height="1"><img src="images/spacer.gif" width="1" height="1"></td>
          </tr>
        </table>
        <table width="98%" border="0" align="center">
          <tr> 
            <td>����λ�ã�<a href="index.asp">��ҳ >></a> 
              <%     
	response.write "��ѯ������<font color=red>"&keyword&"</font>"     
	response.write "��"&classname&""&Nclassname&"��"     
%>
            </td>
          </tr>
        </table>
        <br>
        <TABLE border="0" cellspacing="1" cellpadding="1" bgcolor="#999999" width="98%" align="center">
          <TR>     
              
<TD valign="top" align="center" bgcolor="#FFFFFF">     
                
<TABLE border="0" width="98%" cellspacing="0" cellpadding="0" align="center">    
                  
<TR bgcolor="#FFFFFF">     
                    
<TD bgcolor="#FFFFFF" height="22"> <BR>    
                </TD>    
<TD align=right width="28%" height="22" nowrap>     
                      
<FORM action="sort.asp" method=get>    
                    �ܷ���:       
                          
                      <SELECT name="classid" size="1"  onChange='javascript:submit()' class="textfield">
                        <%      
	sql="select class,classid from class"      
	rs.open sql,conn,1,1      
	if rs.bof and rs.eof then      
%>
                        <OPTION value="">Not Record!</OPTION>      
                            
  <%else%>      
                            
  <%      
	do while not rs.eof      
%>      
                            
  <OPTION <%if request("classid")<>"" then%><%if cint(rs("classid"))=cint(request("classid")) then%> selected  <%end if%><%end if%>value="<%=rs("classid")%>"><%=rs("class")%></OPTION>      
                            
  <%      
	rs.movenext      
	loop      
	end if      
	rs.close      
%>      
                          
 </SELECT>      
                        
</FORM>      
</TD>      
<TD width="29%" height="22" nowrap>       
                        
<FORM action="sort.asp" method=get>      
�ӷ���: 
                      <SELECT name="Nclassid" size="1" onChange='javascript:submit()' class="textfield">
                        <%      
	if request("classid")="" then      
	sql="select Nclass,Nclassid from Nclass"      
	else      
	sql="select Nclass,Nclassid from Nclass where classid="&request("classid")      
	end if      
	rs.open sql,conn,1,1      
	if rs.bof and rs.eof then      
%>
                        <OPTION value="">Not Record!</OPTION>      
                            
  <%else%>      
                            
  <%      
	do while not rs.eof      
%>      
                            
  <OPTION <%if request("Nclassid")<>"" then%><%if cint(rs("Nclassid"))=cint(request("Nclassid")) then%> selected  <%end if%><%end if%>value="<%=rs("Nclassid")%>"><%=rs("Nclass")%></OPTION>      
                            
  <%      
	rs.movenext      
	loop      
	end if      
	rs.close      
%>      
                          
 </SELECT>      
                        
</FORM>      
</TD>      
</TR>      
                  
</TABLE>      
<%       
	if request("Nclassid")="" then      
		sql="select id,showname,bb,note,hits,dateandtime,hot,size,system,orders from download where "&classid&" "&findword&" "      
		sql=sql&" order by id desc"      
	else      
		sql="select id,showname,bb,note,hits,dateandtime,hot,size,system,orders from download where "&classid&" "&Nclassid&" "&findword&" "      
		sql=sql&" order by id desc"      
	end if      
	rs.open sql,conn,1,1      
      
  	if rs.eof and rs.bof then       
       		response.write "<p align='center'>û�л�û���ҵ��κ���Ϣ<br><br><br><b>��ʾ:����ӰƬʱ�벻Ҫ����汾��!</b><br><Br></p>"       
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
            		showpage totalput,MaxPerPage,"search.asp"       
       		else       
          		if (currentPage-1)*MaxPerPage<totalPut then       
            			rs.move  (currentPage-1)*MaxPerPage       
            			dim bookmark       
            			bookmark=rs.bookmark       
            			showContent       
             			showpage totalput,MaxPerPage,"search.asp"       
        		else       
	        		currentPage=1       
           			showContent       
           			showpage totalput,MaxPerPage,"search.asp"       
	      		end if       
	   	end if       
   	rs.close       
   	end if       
	               
   	sub showContent       
       	dim i       
	   	i=0       
%>      
                
                    
<TABLE border="0" cellspacing="0" cellpadding="0" align=center width="98%">      
                      
<TR>       
                        
<TD width="37%" height=22>ӰƬ���ƺͼ��</TD>      
<TD width="18%" align="center">�Ƽ�</TD>      
<TD width="18%" align="center">��������</TD>      
<TD width="13%" align="center">���ش���</TD>      
<TD width="14%" align="center">�ļ���С</TD>      
</TR>      
                      
<%do while not rs.eof%>      
                      
<TR>       
                        
<TD width="100%" colspan="5" height=23><A href="list.asp?id=<%=rs("id")%>"><%=replace(rs("showname"),""&keyword&"","<font color=red>"&keyword&"</font>")%>&nbsp;<%=rs("bb")%></A></TD>      
</TR>      
                      
<TR>       
                        
<TD width="37%" height=22>&nbsp;&nbsp;       
                          
<%if len(rs("note"))>18 then%>      
                    <%=left(rs("note"),18)%>����       
                          
<%else%>      
                    <%=rs("note")%>       
                          
<%end if%>      
                  </TD>      
<TD width="18%" align=center>       
                          
<%if rs("hot")>3 then%>      
                    <FONT color=red>�Ƽ�</FONT>       
                          
<%end if%>      
                  </TD>      
<TD width="18%" align=center><%=rs("dateandtime")%></TD>      
<TD width="13%" align=center><%=rs("hits")%></TD>      
<TD width="14%" align=center><%=rs("size")%></TD>      
</TR>      
                      
<TR>       
                        
<TD width="100%" colspan="5" height=11>���л�����<%=rs("system")%>       
                    ��Ȩ��ʽ��<%=rs("orders")%></TD>      
</TR>      
                      
<TR align="center">       
                        
<TD width="100%" colspan="5"><hr size="1"></TD>      
</TR>      
<%      
	 i=i+1      
	 if i>=MaxPerPage then exit do      
	 	rs.movenext      
	 loop      
%>      
                    
</TABLE>      
              <%      
   end sub       
      
	function showpage(totalnumber,maxperpage,filename)      
  	dim n      
      
  	if totalnumber mod maxperpage=0 then      
     		n= totalnumber \ maxperpage      
  	else      
     		n= totalnumber \ maxperpage+1      
  	end if      
  	response.write "<form method=Post action="&filename&"?classid="&request("classid")&"&Nclassid="&request("Nclassid")&"&action="&request("action")&"&keyword="&keyword&">"      
  	response.write "<font color='red'>"&Nclassname&"</font>"      
      
  	if CurrentPage<2 then      
    		response.write ""&totalnumber&"��&nbsp;��ҳ ��һҳ&nbsp;"      
  	else      
    		response.write ""&totalnumber&"��&nbsp;<a href="&filename&"?page=1&keyword="&keyword&">��ҳ</a>&nbsp;"      
    		response.write "<a href="&filename&"?page="&CurrentPage-1&"&keyword="&keyword&">��һҳ</a>&nbsp;"      
  	end if      
      
  	if n-currentpage<1 then      
    		response.write "��һҳ βҳ"      
  	else      
    		response.write "<a href="&filename&"?page="&(CurrentPage+1)&"&keyword="&keyword&">"      
    		response.write "��һҳ</a> <a href="&filename&"?page="&n&"&keyword="&keyword&">βҳ</a>"      
  	end if      
   	response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&CurrentPage&"</font>/"&n&"</strong>ҳ "      
    	response.write "&nbsp;<b>"&maxperpage&"</b>��ӰƬ/ҳ "      
%>
   <%           
end function      
%>
              <%      
   	set rs=nothing         
	conn.close      
	set conn=nothing      
%>
            </TD>      
</TR>      
            
</TABLE>      
</div><br>   
</td>      
<%end if%>      
<!--#include file="CopyRight.asp"-->      
</table>    
</BODY>      
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
 
 
 
 
 
 
 
 
 
 
 
 
 






