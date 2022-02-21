<%
if request.cookies("adminok")="" then
  response.redirect "login.asp"
end if
%>
<!--#include file="articleconn.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>管理文件</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<link rel="stylesheet" href="css/article.css">
</head>
<%
   const MaxPerPage=20 
   dim totalPut   
   dim CurrentPage
   dim TotalPages
   dim i,j

   if not isempty(request("page")) then
      currentPage=cint(request("page"))
   else
      currentPage=1
   end if
   
%>
<body bgcolor="#FFFFFF">
<p>&nbsp;</p>
<table width="700" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="10" bgcolor="#B5D85E"> 
      <div align="center"><b>管 理 界 面</b></div>
    </td>
  </tr>
  <tr> 
    <td height="49"><%
dim sql
dim rs
sql="select * from learning order by articleid desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
  if rs.eof and rs.bof then
       response.write "<p align='center'> 还 没 有 任 何 信 息</p>"
   else
	  totalPut=rs.recordcount
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
                    showpages
                    showContent
					showpages
	           else
			       if (currentPage-1)*MaxPerPage<totalPut then
		            rs.move  (currentPage-1)*MaxPerPage
		            dim bookmark
		            bookmark=rs.bookmark
                    showpages
                    showContent
					showpages
                  else
			        currentPage=1
                   showpages
                    showContent
                   showpages
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
      <table border="1" cellspacing="0" width="90%" bgcolor="#F0F8FF" bordercolorlight="#000000"
bordercolordark="#FFFFFF" align="center">
        <tr> 
          <td width="12%" align="center"><strong>ID 号</strong></td>
          <td width="13%" align="center"><b>类 型</b></td>
          <td width="47%" align="center"><strong>信 息 名 称</strong></td>
          <td width="14%" align="center"><strong>修 改</strong></td>
          <td width="14%" align="center"><strong>删 除</strong></td>
        </tr>
        <%do while not rs.eof%> 
        <tr> 
          <td width="12%" height="7"> 
            <p align="center"><%=rs("articleid")%> 
          </td>
          <td width="13%" height="7"> 
            <div align="center"><%=rs("type")%></div>
          </td>
          <td width="47%" height="7"><a href="viewarticle.asp?id=<%=rs("articleid")%>"><%=rs("title")%></a></td>
          <td width="14%" align="center" height="7"><a
    href="edit.asp?id=<%=rs("articleid")%>">修 改</a></td>
          <td width="14%" align="center" height="7"><a
    href="delete.asp?id=<%=rs("articleid")%>">删 除</a></td>
        </tr>
        <% i=i+1
	      if i>=MaxPerPage then exit do
	      rs.movenext
	   loop
		  %> 
      </table>
      <p><%
   end sub 

   sub showpages()
          dim n
	   if (totalPut mod MaxPerPage)=0 then
	      n= totalPut \ MaxPerPage
	   else
	      n= totalPut \ MaxPerPage + 1
	   end if
	   if n=1 then 
        response.write "<p align='left'><a href=add.asp>添加信息</a>"
	     response.write "</p>"
        exit sub
       end if

	   dim k
	   response.write "<p align='left'>&gt;&gt; 信息分页 "
	   for k=1 to n
	       if k=currentPage then
	          response.write "[<b>"+Cstr(k)+"</b>] "
		   else
		      response.write "[<b>"+"<a href='manage.asp?page="+cstr(k)+"'>"+Cstr(k)+"</a></b>] "
		   end if
	   next
      response.write " <a href=add.asp>创建信息</a>"
	   response.write "</p>"
   end sub

 
%></p>
    </td>
  </tr>
</table>
<div align="center">
  <center>
  </center>
</div>

</body>
</html>
