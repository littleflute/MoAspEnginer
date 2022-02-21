<%
  	if session("admin")="" then
  		response.redirect "admin.asp"
  		response.end
  	end if
%>
<!--#include file="conn.asp"-->
<!--#include file="../inc/char.asp"-->
<!--#include file="../inc/const.asp"-->
<%
	dim filename,filename1,filename2,filename3,filename4
	dim txtname,txtname1,txtname2,txtname3,txtname4
	dim showname
	dim bb
	dim classid
	dim Nclassid
	dim note
	dim sql
	dim rs
	dim hot
	dim system
	dim fromurl
	dim size
        dim images
	dim order
        dim xxyf
        dim hy
	dim hots
	dim hide
	dim founderr
	founerr=false
	if trim(request.form("txtshowname"))="" then
  		founderr=true
  		errmsg=errmsg+"<li>软件名称不能为空</li>"
	end if
	if trim(request.form("txtnote"))="" then
  		founderr=true
  		errmsg=errmsg+"<li>软件简介不能为空</li>"
	end if

	if founderr=false then
	    movie=request("movie")
		filename=request("txtfilename")
		filename1=request("txtfilename1")
		filename2=request("txtfilename2")
		filename3=request("txtfilename3")
		filename4=request("txtfilename4")
		txtname=request("txtname")
		txtname1=request("txtname1")
		txtname2=request("txtname2")
		txtname3=request("txtname3")
		txtname4=request("txtname4")
		showname=htmlencode2(request.form("txtshowname"))
		bb=htmlencode2(request("txtbb"))
		classid=request("classid")
		Nclassid=request("Nclassid")
		note=htmlencode2(request("txtnote"))
		hot=htmlencode2(request("hot"))
		system=htmlencode2(request("system"))
		size=htmlencode2(request("size"))
                images=request("images")
		fromurl=htmlencode2(request("fromurl"))
		order=htmlencode2(request("order"))
                xxyf=htmlencode2(request("xxyf"))
                hy=htmlencode2(request("hy"))
		if request.form("hots")="on" then
			hots=1
		else
			hots=0
		end if
		if request.form("hide")="on" then
			hide=1
		else
			hide=0
		end if

	set rs=server.createobject("adodb.recordset")
	if request("action")="add" then
		call newsoft()
	elseif request("action")="edit" then
		call editsoft()
	end if
sub newsoft()
	sql="select * from download where (id is null)" 
	rs.open sql,conn,1,3
	rs.addnew
	rs("movie")=movie
	
	rs("txtname")=txtname
	if txtname1<>"" then
	rs("txtname1")=txtname1
	end if
	if txtname2<>"" then
	rs("txtname2")=txtname2
	end if
	if txtname3<>"" then
	rs("txtname3")=txtname3
	end if
	if txtname4<>"" then
	rs("txtname4")=txtname4
	end if
	
	rs("filename")=filename
	if filename1<>"" then
	rs("filename1")=filename1
	end if
	if filename2<>"" then
	rs("filename2")=filename2
	end if
	if filename3<>"" then
	rs("filename3")=filename3
	end if
	if filename4<>"" then
	rs("filename4")=filename4
	end if
	rs("showname")=showname
	rs("bb")=bb
	rs("classid")=classid
	rs("Nclassid")=Nclassid
	rs("fromurl")=fromurl
	rs("note")=note
	rs("system")=system
	rs("hot")=hot
	rs("hots")=hots
	rs("stop")=hide
	if size<>"" and size<>"K" then
	rs("size")=size
	end if
    rs("images")=images
	rs("orders")=order
        rs("xxyf")=xxyf
        rs("hy")=hy
	rs("lasthits")=date()
	rs("dateandtime")=now()
	rs("dayhits")=0
	rs("weekhits")=0
	rs("hits")=0
	rs.update
end sub
sub editsoft()
	sql="select * from download where id="&request("id")
	rs.open sql,conn,1,3
	rs("movie")=movie
	rs("filename")=filename
	rs("filename1")=filename1
	rs("filename2")=filename2
	rs("filename3")=filename3
	rs("filename4")=filename4
	rs("txtname")=txtname
	rs("txtname1")=txtname1
	rs("txtname2")=txtname2
	rs("txtname3")=txtname3
	rs("txtname4")=txtname4
	rs("showname")=showname
	rs("bb")=bb
	rs("classid")=classid
	rs("Nclassid")=Nclassid
	rs("fromurl")=fromurl
	rs("note")=note
	rs("system")=system
	rs("hot")=hot
	rs("hots")=hots
	rs("stop")=hide
	if size="" or size="K" then
	rs("size")=Null
	else
	rs("size")=size
	end if
    rs("images")=images
	rs("orders")=order
        rs("xxyf")=xxyf
        rs("hy")=hy
	rs("lasthits")=date()
	rs("dateandtime")=now()
	rs.update
end sub
	set rs=nothing
	conn.close
	set conn=nothing
%>
<html>

<head>
<title></title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body bgcolor="#39867B">
<br><br>
  
<table width="50%" border="0" cellpadding="3" cellspacing="1" align="center">
  <tr>
    
<td width="100%" height="20" bgcolor="#E1F4EE">
<p align="center"><FONT color="#000000">
 <%if request("action")="add" then%>
 添加
 <%else%>
 修改
 <%end if%>
 程序成功</font>
</td>
  </tr>
  <tr>
    
<td width="100%" bgcolor="#E1F4EE">
<p align="left"><br>
    显示名称为：<%=request.form("txtshowname")%>&nbsp;<%=bb%><br>
    下载地址1为：<%=filename%> 说明是：<%=txtname%><br>
    下载地址2为：<%=filename1%> 说明是：<%=txtname1%><br>
    下载地址3为：<%=filename2%> 说明是：<%=txtname2%><br>
    下载地址4为：<%=filename3%> 说明是：<%=txtname3%><br>
    下载地址5为：<%=filename4%> 说明是：<%=txtname4%><br>

    </p>
    您可以进行其他操作 <a href="freeadd.asp" target="_self">添加</a> <a href="adminedit.asp" target="_self">修改删除</a></td>
  </tr>
</table>
<%
else
 response.write "由于以下的原因不能保存数据："
 response.write errmsg
end if
%>

</body>
</html>