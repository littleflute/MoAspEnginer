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
		classname="�������"
	else
		classid="classid="&cstr(request("classid"))&" and  "
		sql="select class from class where classid="&cstr(request("classid"))
		rs.open sql,conn,1,1
		classname="[<font color=#008000>"&rs("class")&"</font>]"
		rs.close
	end if
	if request("Nclassid")="" then
		Nclassid=""
		Nclassname="�������"
	else
		Nclassid=" Nclassid="&cstr(request("Nclassid"))&" and  "
		sql="select Nclass from Nclass where Nclassid="&cstr(request("Nclassid"))
		rs.open sql,conn,1,1
		Nclassname="��[<font color=#008000>"&rs("Nclass")&"</font>]����"
		rs.close
	end if

%>
<HTML>
<HEAD>
<TITLE>���ҽ��(<%=keyword%>)-</TITLE>
<META content="text/html; charset=gb2312" http-equiv="Content-Type">
</HEAD>
<!--#include file="head.asp"-->
<!--#include file="lanmu.asp"-->
<br>
<%if founderr=true then%>
<TABLE border="0" width="770" cellpadding="0" cellspacing="1" align="center">
   
  <TR> 
    <TD width="168" valign="top"> <table border=0 cellpadding="0" cellspacing="0"  width="168" height="73">
        <tr> 
          <td width="174" height="22" class="li"> ��������</td>
        </tr>
        <tr bgcolor=#ffffff> 
          <td width="174" align="center" class="3b" background="pic/a-27.gif" bgcolor="#ffffff"><TABLE width="98%" height="63" border=0 align=center cellPadding=0 cellSpacing=2>
              <FORM action=query.asp method=POST name="myfrom">
                <TR> 
                  <TD> <input onfocus="this.value=''" maxLength="50" size="17" title="����ؼ���" value="����ؼ���" name="keyword2" type="text"> 
                  </TD>
                </TR>
                <TR> 
                  <TD><input type="radio" value="title" checked name="action">
                    ���� 
                    <input type="radio" name="action" value="content">
                    ���<font color="#FFFFFF">&nbsp;</font></TD>
                </TR>
                <TR> 
                  <TD><input type=submit value="�� ��" name=Submit> <a href="search.asp"> 
                    �߼�����...</a></TD>
                </TR>
              </FORM>
            </TABLE></td>
      </table>
      </TD>
    <TD width="558" bgcolor="#FFFFFF" valign="top"> 
      <TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
         
        <TR> 
          <TD height="5"></TD>
        </TR>
        <TR> 
          <TD background="images/bj4.gif" height="1"><IMG src="images/spacer.gif" width="1" height="1"></TD>
        </TR>
         
      </TABLE>
      <TABLE width="98%" border="0" align="center">
         
        <TR> 
          <TD>����λ�ã�<a href="<%=rs1("urls")%>"><%=rs1("homes")%></a>&gt;&gt; 
            <%     
	response.write "��ѯ������<font color=red>"&keyword&"</font>"     
	response.write "��"&classname&""&Nclassname&"��"     
%>
          </TD>
        </TR>
         
      </TABLE>
      <P><BR>
        <BR>
        <BR>
        <BR>
        <BR>
        <BR>
        <BR>
      </P>
      <DIV align="center"> �ؼ��ֲ���Ϊ��!</DIV>
    </TD>
  </TR>
   
</TABLE>
<%else%>
<table border="0" bgcolor="#FFFFFF" cellpadding="0" cellspacing="0"  width="770" id="AutoNumber1">
  <tr> 
    <td width="168" valign="top"> <table border=0 cellpadding="0" cellspacing="0"  width="168" height="73">
        <tr> 
          <td width="174" height="22" class="li"> ��������</td>
        </tr>
        <tr bgcolor=#ffffff> 
          <td width="174" align="center" class="3b" background="pic/a-27.gif" bgcolor="#ffffff"><TABLE width="98%" height="63" border=0 align=center cellPadding=0 cellSpacing=2>
              <FORM action=query.asp method=POST name="myfrom">
                <TR> 
                  <TD> <input onfocus="this.value=''" maxLength="50" size="17" title="����ؼ���" value="����ؼ���" name="keyword" type="text"> 
                  </TD>
                </TR>
                <TR> 
                  <TD><input type="radio" value="title" checked name="action">
                    ���� 
                    <input type="radio" name="action" value="content">
                    ���<font color="#FFFFFF">&nbsp;</font></TD>
                </TR>
                <TR> 
                  <TD><input type=submit value="�� ��" name=Submit2> <a href="search.asp"> 
                    �߼�����...</a></TD>
                </TR>
              </FORM>
            </TABLE></td>
      </table>
      <br>
      <table border=0 cellpadding="0" cellspacing="0"  width="168">
        <tr> 
          <td width="181" class="li" height="22"> ������������</td>
        </tr>
        <tr> 
          <td width="181" class="3b" background="pic/a-27.gif"> 
            <%
	dim tdate
	tdate=year(Now()) & "-" & month(Now()) & "-" & day(Now())
	sql="select top 10 id,showname,bb,dayhits from download where "
	sql=sql&" lasthits="&tdate&" and dayhits>0 "
	sql=sql&" order by dayhits desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
response.write ("����û������")
else
do while not rs.eof
response.write ("��<A href=software.asp?id="&rs("id")&" title=�������أ�"&rs("dayhits")&"��>"&rs("showname")&"</A><br>")
i=i+1
if i>=10 then exit do
rs.movenext
loop
i="0"
end if
rs.close
%>
          </td>
        </tr>
      </table>
      <br>
      <table border=0 cellpadding="0" cellspacing="0"  width="168">
        <tr> 
          <td class="li" height="22">������������</td>
        </tr>
        <tr> 
          <td width="181" class="3b" background="pic/a-27.gif"> 
            <%                                                                                     
	OldWeek = WeekDay(Date())-1                                                                                     
	If OldWeek = 0 Then OldWeek = 7                                                                                     
	OldWeek = Date()-OldWeek                                                                                     
	NewWeek = Date()+(9-WeekDay(Date()))                                                                                     
	sql="select top 10 id,showname,bb from download where "                                                                                     
	sql=sql&"  (lasthits < " & NewWeek & ") And (lasthits > " & OldWeek & ") and weekhits>0 "                                                                                     
	sql=sql&" order by weekhits desc"                                                                                     
	rs.open sql,conn,1,1                                                                                     
	if rs.eof and rs.bof then                                                                                     
%>
            <p>����û������ </p>
            <%else%>
            <%do while not rs.eof                                                                                     
response.write "��<A href=software.asp?id="&rs("id")&">"&rs("showname")&"</A><br>"                                                                                     
i=i+1                                                
if i>=10 then exit do                                                
rs.movenext                                                
loop                                                
i="0"                                                
	end if                                                 
	rs.close                                                 
%>
        </tr>
      </table>
      <br>
      <table border=0 cellpadding="0" cellspacing="0"  width="168">
        <tr> 
          <td width="181" height="22" class="li">ȫ����������</td>
        </tr>
        <tr> 
          <td width="181" class="3b" background="pic/a-27.gif"> 
            <%                   
	sql="select top 10 id,showname,bb,hits,dayhits from download "                   
	sql=sql&" order by hits desc"                   
	rs.open sql,conn,1,1                   
	if rs.eof and rs.bof then
response.write ("û������")
else
do while not rs.eof
response.write ("��<A href=software.asp?id="&rs("id")&" title=�����أ�"&rs("hits")*3&"��&#13;&#10;�������أ�"&rs("dayhits")&"��>"&rs("showname")&"</A><br>")
	rs.movenext                                     
	loop                                     
	end if                                     
	rs.close                                     
%>
        </tr>
      </table>
    <td align="middle" valign="top" width="560" bgcolor="#FFFFFF"> 
      <table border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF"  width="560">
        <tr> 
          <td align="center" width="560" bgcolor="#FFFFFF" valign="top"> 
            <table border="0" width="560" cellspacing="0" cellpadding="0" height="22"  bordercolor="#111111">
              <tr> 
                <td bgcolor="#FFFFFF" height="22" valign="top"> <a href="<%=rs1("urls")%>"><%=rs1("homes")%></a> 
                  >> 
                  <%     
	response.write "��ѯ������<font color=red>"&keyword&"</font>"     
	response.write "��"&classname&""&Nclassname&"��"     
%>
                </td>
              </TR>
            </TABLE>
            <P><BR>
            </P>
            <TABLE border="0" cellspacing="0" cellpadding="0" bgcolor="#999999" width="98%" align="center">
              <TR> 
                <TD valign="top" align="center" bgcolor="#FFFFFF"> 
                  <TABLE border="0" width="98%" cellspacing="0" cellpadding="0" align="center">
                     
                    <TR bgcolor="#FFFFFF"> 
                      <TD bgcolor="#FFFFFF" height="22"><BR>
                      </TD>
                      <TD align="right" width="28%" height="22" nowrap> 
                        <FORM action="class.asp" method="get">
                          <P>�ܷ���: 
                            <SELECT name="classid" size="1" onChange="javascript:submit()" class="textfield">
                              <%      
	sql="select class,classid from class"      
	rs.open sql,conn,1,1      
	if rs.bof and rs.eof then      
%>
                              <OPTION value>Not Record!</OPTION>
                              <%else%>
                              <%      
	do while not rs.eof      
%>
                              <OPTION <%if request("classid")<>"" then%><%if cint(rs("classid"))=cint(request("classid")) then%> selected <%end if%><%end if%>value="<%=rs("classid")%>"><%=rs("class")%></OPTION>
                              <%      
	rs.movenext      
	loop      
	end if      
	rs.close      
%>
                            </SELECT>
                          </P>
                        </FORM>
                      </TD>
                      <TD width="29%" height="22" nowrap> 
                        <FORM action="class.asp" method="get">
                          <P>�ӷ���: 
                            <SELECT name="Nclassid" size="1" onChange="javascript:submit()" class="textfield">
                              <%      
	if request("classid")="" then      
	sql="select Nclass,Nclassid from Nclass"      
	else      
	sql="select Nclass,Nclassid from Nclass where classid="&request("classid")      
	end if      
	rs.open sql,conn,1,1      
	if rs.bof and rs.eof then      
%>
                              <OPTION value>Not Record!</OPTION>
                              <%else%>
                              <%      
	do while not rs.eof      
%>
                              <OPTION <%if request("Nclassid")<>"" then%><%if cint(rs("Nclassid"))=cint(request("Nclassid")) then%> selected <%end if%><%end if%>value="<%=rs("Nclassid")%>"><%=rs("Nclass")%></OPTION>
                              <%      
	rs.movenext      
	loop      
	end if      
	rs.close      
%>
                            </SELECT>
                          </P>
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
       		response.write "<p align='center'>SORRY!!û���ҵ���Ҫ�ĳ����볢��ʹ���������ƵĹؼ��֣�<br>����ȥ��̳�����򷢱�</p>"       
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
            		showpage totalput,MaxPerPage,"query.asp"       
       		else       
          		if (currentPage-1)*MaxPerPage<totalPut then       
            			rs.move  (currentPage-1)*MaxPerPage       
            			dim bookmark       
            			bookmark=rs.bookmark       
            			showContent       
             			showpage totalput,MaxPerPage,"query.asp"       
        		else       
	        		currentPage=1       
           			showContent       
           			showpage totalput,MaxPerPage,"query.asp"       
	      		end if       
	   	end if       
   	rs.close       
   	end if       
	               
   	sub showContent       
       	dim i       
	   	i=0       
%>
                  <TABLE border="0" width="540" cellspacing="0" cellpadding="0"  bordercolor="#111111">
                    �������ǲ�ѯ���: 
                    <TR> 
                      <TD height=20 bgcolor="#F3F3F3" width="236"><b>������ƺͼ��</b></TD>
                      <TD align="center" bgcolor="#F3F3F3" width="67"><b>�Ƽ�</b></TD>
                      <TD align="center" bgcolor="#F3F3F3" width="99"><b>��������</b></TD>
                      <TD align="center" bgcolor="#F3F3F3" width="71"><b>���ش���</b></TD>
                      <TD align="center" bgcolor="#F3F3F3" width="77"><b>�ļ���С</b></TD>
                    </TR>
                    <%do while not rs.eof%>
                    <TR> 
                      <TD width="236" colspan="5" height="23"><IMG SRC="./images/list.gif" WIDTH=9 HEIGHT=17 ALT="" BORDER=0>&nbsp;<A href="software.asp?id=<%=rs("id")%>"><b><%=replace(rs("showname"),""&keyword&"","<font color=red>"&keyword&"</font>")%></b></A></TD>
                    </TR>
                    <TR> 
                      <TD width="37%" height="22">&nbsp;&nbsp; 
                        <%if len(rs("note"))>18 then%>
                        <%=left(rs("note"),18)%>���� 
                        <%else%>
                        <%=rs("note")%> 
                        <%end if%>
                      </TD>
                      <TD width="67" align="center"> 
                        <%if rs("hot")>3 then%>
                        <FONT color="red">�Ƽ�</FONT> 
                        <%end if%>
                      </TD>
                      <TD width="18%" align="center"><%=rs("dateandtime")%></TD>
                      <TD width="13%" align="center"><%=rs("hits")%></TD>
                      <TD width="14%" align="center"><%=rs("size")%></TD>
                    </TR>
                    <TR> 
                      <TD width="100%" colspan="5" height="11">���л�����<%=rs("system")%> 
                        ��Ȩ��ʽ��<%=rs("orders")%></TD>
                    </TR>
                    <TR align="center"> 
                      <TD width="100%" colspan="5"> 
                        <HR size="1">
                      </TD>
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
    	response.write "&nbsp;<b>"&maxperpage&"</b>�����/ҳ "      
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
          </TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
</TABLE>
<%end if%>
<!--#include file="foot.asp"-->
</BODY>
</HTML>