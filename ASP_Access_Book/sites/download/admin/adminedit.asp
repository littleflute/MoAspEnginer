<%
  if session("admin")="" then
  response.redirect "admin.asp"
  else
	if session("flag")>2 then
		response.write "<br><p align=center>您没有操作的权限</p>"
		response.end
	end if
  end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<!--#include file="conn.asp"-->
<%
   	const MaxPerPage=18
   	dim totalPut   
   	dim CurrentPage
   	dim TotalPages
   	dim i,j
   	dim idlist
   	dim title

	title=request("txtitle")
   	if not isempty(request("page")) then
      		currentPage=cint(request("page"))
   	else
      		currentPage=1
   	end if
  	if not isempty(request("selAnnounce")) then
     		idlist=request("selAnnounce")
     		if instr(idlist,",")>0 then
			dim idarr
			idArr=split(idlist)
			dim id
		for i = 0 to ubound(idarr)
	       		id=clng(idarr(i))
	       		call deleteannounce(id)
		next
     		else
			call deleteannounce(clng(idlist))
     		end if
  	end if 
   	dim sql
   	dim rs
%>
<body bgcolor="#39867B">
<table border="0" width="92%" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="100%" align="center"> <FONT color="#FF9933">点击相关软件察看及修改软件资料，删除软件请选中相关软件然后点击删除 
      </FONT> 
      <form name="searchsoft" method="POST" action="adminedit.asp">
        查找软件: 
        <input class=smallInput type="text" name="txtitle" size="13">
        <input class=buttonface type="submit" value="查 询" name="title">
        &nbsp; 排列顺序： 
        <%if request("pl")="n" then%>
        <b> 
        <%end if%>
        </b><a href="adminedit.asp?pl=n"><b>软件名</b></a>、<a href="adminedit.asp?pl=d"> 
        <%if request("pl")="d" or request("pl")="dateandtime" then%>
        <b> 
        <%end if%>
        日期</b></a>、<a href="adminedit.asp?pl=i"> 
        <%if request("pl")="i" then%>
        <b> 
        <%end if%>
        ID</b></a> 
      </form>
      <form method=Post action="adminedit.asp">
        <table border="0" cellspacing="1" width="589" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="2">
          <tr bgcolor="#E1F4EE"> 
            <td width="46" align="center" height="20"><STRONG>ID</STRONG></td>
            <td width="400" align="center"><STRONG>软 件 名</STRONG></td>
            <td width="68" align="center"> 
              <input type='submit' class=buttonface value='删 除'>
            </td>
          </tr>
          <tr bgcolor="#E1F4EE"> 
            <td colspan="3" align="center" height="20"> 
              <%
    if request("pl")="n" then
    ppp="showname"
    elseif request("pl")="i" then
    ppp="id"
    else
    ppp="dateandtime"
    end if
	if title<>"" then 
	sql="select * from download where showname like '%"&trim(title)&"%' order by "&ppp&" desc" 
	else 
	sql="select * from download order by "&ppp&" desc" 
	end if 
	Set rs= Server.CreateObject("ADODB.Recordset") 
	rs.open sql,conn,1,1 
 
  	if rs.eof and rs.bof then 
       		response.write "<p align='center'> 还 没 有 任 何 软 件 </p>" 
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
           		showpage totalput,MaxPerPage,"adminedit.asp" 
            		showContent 
            		showpage totalput,MaxPerPage,"adminedit.asp" 
       		else 
          		if (currentPage-1)*MaxPerPage<totalPut then 
            			rs.move  (currentPage-1)*MaxPerPage 
            			dim bookmark 
            			bookmark=rs.bookmark 
           			showpage totalput,MaxPerPage,"adminedit.asp" 
            			showContent 
             			showpage totalput,MaxPerPage,"adminedit.asp" 
        		else 
	        		currentPage=1 
           			showpage totalput,MaxPerPage,"adminedit.asp" 
           			showContent 
           			showpage totalput,MaxPerPage,"adminedit.asp" 
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
            </td>
          </tr>
          <%do while not rs.eof%>
          <tr> 
            <td height="11" width="46" bgcolor="#E1F4EE"> 
              <p align="center"><%=rs("id")%> 
            </td>
            <td width="400" bgcolor="#E1F4EE"> 
              <p><a href="editsoft.asp?id=<%=rs("id")%>&classid=<%=rs("classid")%>&Nclassid=<%=rs("Nclassid")%>"><%=rs("showname")%></a>
            </td>
            <td width="68" bgcolor="#E1F4EE"> 
              <p align="center"> 
                <input type='checkbox' name='selAnnounce' value='<%=cstr(rs("ID"))%>'>
            </td>
          </tr>
          <% 
	i=i+1 
	      if i>=MaxPerPage then exit do 
	      rs.movenext 
	loop 
%>
          <tr bgcolor="#E1F4EE" align="center"> 
            <td height="11" colspan="3"> 
              <DIV align="center"> 
                <CENTER>
                </CENTER>
              </DIV>
              <% 
   end sub  
 
function showpage(totalnumber,maxperpage,filename) 
  dim n 
  if totalnumber mod maxperpage=0 then 
     n= totalnumber \ maxperpage 
  else 
     n= totalnumber \ maxperpage+1 
  end if 
  response.write "<p align='center'>&nbsp;" 
  if CurrentPage<2 then 
    response.write "<font color='#000080'>首页 上一页</font>&nbsp;" 
  else 
    response.write "<a href="&filename&"?page=1&pl="&request("pl")&">首页</a>&nbsp;" 
    response.write "<a href="&filename&"?page="&CurrentPage-1&"&pl="&request("pl")&">上一页</a>&nbsp;" 
  end if 
  if n-currentpage<1 then 
    response.write "<font color='#000080'>下一页 尾页</font>" 
  else 
    response.write "<a href="&filename&"?page="&(CurrentPage+1)&"&pl="&request("pl")&">" 
    response.write "下一页</a> <a href="&filename&"?page="&n&"&pl="&request("pl")&">尾页</a>" 
  end if 
   response.write "<font color='#000080'>&nbsp;页次：</font><strong><font color=red>"&CurrentPage&"</font><font color='#000080'>/"&n&"</strong>页</font> " 
    response.write "<font color='#000080'>&nbsp;共<b>"&totalnumber&"</b>个软件 <b>"&maxperpage&"</b>个软件/页</font> " 
        
end function 
 
  sub deleteannounce(id) 
    dim rs,sql 
    set rs=server.createobject("adodb.recordset") 
    sql="delete from download where id="&cstr(id) 
    conn.execute sql 
    if err.Number<>0 then 
	err.clear 
	response.write "删 除 失 败 !<br>" 
    else 
        response.write "操作成功！<br>" 
    end if 
  End sub 
%>
            </td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
</body>
</html> 
