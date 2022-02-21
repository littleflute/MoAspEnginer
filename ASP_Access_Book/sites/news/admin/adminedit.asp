<!--#include file="conn.asp"-->
<%
  if session("admin")="" then
  response.redirect "admin.asp"
  else
	if session("flag")>2 then
		response.write "<br><p align=center>您没有操作的权限</p>"
		response.end
	end if
  end if
dim action,sql,rs
action=request("action")
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body>
<%if not isempty(request("selAnnounce")) then
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

if action="" then
sql="select * from class order by rootID,ordersID"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
rs.close
response.write "还没有相应的分类。"
else%>
<div align="center">
  <center>
<table width="95%" cellspacing="0" cellpadding="0" class="border" style="border-collapse: collapse" bordercolor="#111111">
<tr> 
<th width="39%" class="title" height=25>分类
</th>
<th width="31%" class="title" height=25>管理操作</th>
</tr>
<%j=0
do while not rs.eof
%>
<%if j mod 2 = 1 then%><tr class="tdbg"><%else%><tr class="tdbg2"><%end if %>
<td height="25" width="39%">
<%if rs("layer")>0 then%>
<%for i=1 to rs("layer")%>
&nbsp;
<%next%>
<%end if%>
<%if rs("child")>0 then%><img src="pic/plus.gif"><%else%><img src="pic/nofollow.gif"><%end if%>
<%if rs("parentID")=0 then%><b><%end if%><%=rs("class")%><%if rs("child")>0 then%>(<%=rs("child")%>)<%end if%>
</td>
<td width="61%">
<p align="center"><%if rs("classtype")>0 then%><a href="freeadd.asp?action=add&classid=<%=rs("classid")%>&classtype=<%=rs("classtype")%>"><U>
添加内容</U></a> | <a href="adminedit.asp?action=edit&classid=<%=rs("classid")%>&classtype=<%=rs("classtype")%>"><u>
内容管理</u></a> |<%end if%></td>
</tr>
<%
rs.movenext
j=j+1
loop
end if
%>
</table>
  </center>
</div>
<%
rs.close
end if

if action="edit" then
if session("flag")<3 then
classtype=request("classtype")%>
<center>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="98%" class="border">
  <tr>
    <td width="100%" align=center class="title"> <a href="adminedit.asp">返回管理页面</a> | 
    <a href="freeadd.asp?classid=<%=request("classid")%>&classtype=<%=classtype%>">添加内容</a></td>
  </tr>
</table></center>
<%classid=request("classid")
dim title
title=request("txtitle")
if not isempty(request("page")) then
      	currentPage=cint(request("page"))
else
      	currentPage=1
end if
sql="select articleid,classid,title from article"

	if title<>"" then
	sql=sql&" where title like '%"&trim(title)&"%'"
		if classid<>"" then
			sql=sql&" and classid = "&cint(classid)
		end if
	else
		if classid<>"" then
			sql=sql&" where classid = "&cint(classid)
		end if
	end if
	sql=sql&" order by articleid desc"
dim maxperpage
maxperpage=20
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
       		response.write "<p align='center'>暂无数据！ </p>"
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
           		showpage totalput,MaxPerPage,"adminedit.asp?action=edit&classid="&classid&"&classtype="&classtype
            		showContent
            		showpage totalput,MaxPerPage,"adminedit.asp?action=edit&classid="&classid&"&classtype="&classtype
       		else
          		if (currentPage-1)*MaxPerPage<totalPut then
            			rs.move  (currentPage-1)*MaxPerPage
            			dim bookmark
            			bookmark=rs.bookmark
           			showpage totalput,MaxPerPage,"adminedit.asp?action=edit&classid="&classid&"&classtype="&classtype
            			showContent
             			showpage totalput,MaxPerPage,"adminedit.asp?action=edit&classid="&classid&"&classtype="&classtype
        		else
	        		currentPage=1
           			showpage totalput,MaxPerPage,"adminedit.asp?action=edit&classid="&classid&"&classtype="&classtype
           			showContent
           			showpage totalput,MaxPerPage,"adminedit.asp?action=edit&classid="&classid&"&classtype="&classtype
	      		end if
	   	end if
   	rs.close
   	end if
else
response.write "<br><p align=center>您没有操作的权限</p>" 
end if
end if
%>
<div align="center">
		<form name="searchsoft" method="POST" action="adminedit.asp?action=edit">
		查找:  <input class=smallInput type="text" name="txtitle" size="20" class="smallinput">&nbsp;<input type="submit" value="查 询" name="title" class="but"><input type="hidden" name="Nclassid" value="<%=Nclassid%>"><input type="hidden" name="classid" value="<%=classid%>"><input type="hidden" name="classtype" value=<%=classtype%>>
		</form>

<%  
sub showContent
		dim i
		i=0
		%>
      <table class="border" border="0" cellspacing="0" width="95%" cellpadding="3" align="center">
	<form method=Post action="adminedit.asp?action=edit"><input type="hidden" name="classid" value=<%=request("classID")%>><input type="hidden" name="classtype" value=<%=classtype%>><tr class="title">
				  <td width="46" align="center"  height="20"><strong>ID</strong></td>
				  <td width="400" align="center" ><strong>资&nbsp;&nbsp;&nbsp;料</strong></td>
				  <td width="68" align="center" ><input type='submit' value='删 除' class="but"></td>
				</tr>
		<%do while not rs.eof%>
		<%if i mod 2 = 1 then%>
				<tr class="tdbg">
		<%else%>
				<tr class="tdbg2">
		<%end if %>
				  <td height="20" width="46"><p align="center"><%=rs(0)%></td>
				  <td width="400"><a href="editarticle.asp?id=<%=rs(0)%>&classid=<%=rs(1)%>&classtype=<%=request("classtype")%>"><%=rs(2)%></a>
				  </td>
				  <td width="68"><p align="center"><input type='checkbox' name='selAnnounce' value='<%=cstr(rs(0))%>'></td>
				</tr>
		<%
			i=i+1
				  if i>=MaxPerPage then exit do
				  rs.movenext
			loop
		%>
      </form></table>
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
    response.write "<a href="&filename&"&page=1>首页</a>&nbsp;"
    response.write "<a href="&filename&"&page="&CurrentPage-1&">上一页</a>&nbsp;"
  end if
  if n-currentpage<1 then
    response.write "<font color='#000080'>下一页 尾页</font>"
  else
    response.write "<a href="&filename&"&page="&(CurrentPage+1)&">"
    response.write "下一页</a> <a href="&filename&"&page="&n&">尾页</a>"
  end if
   response.write "<font color='#000080'>&nbsp;页次：</font><strong><font color=red>"&CurrentPage&"</font><font color='#000080'>/"&n&"</strong>页</font> "
    response.write "<font color='#000080'>&nbsp;共<b>"&totalnumber&"</b>条 <b>"&maxperpage&"</b>条/页</font> "
end function

sub deleteannounce(id)
    response.write "<center><font color=red>"
    set rs=server.createobject("adodb.recordset")
    sql="delete from article where articleid="&cstr(id)
    conn.execute sql
    if err.Number<>0 then
	err.clear
	response.write "删 除 失 败 !"
    else
    response.write "删除成功！"
    end if
    response.write "</font></center>"
End sub
   	set rs=nothing  
   	conn.close
   	set conn=nothing%>
</body>
</html>