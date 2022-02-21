<!--#include file="conn.asp"-->
<%
  	if session("admin")="" then
  		response.redirect "admin.asp"
  	end if
%>
<!--#include file="char.asp"-->
<%
	dim rs,sql
	dim title,writefrom
	dim addtime
	dim content,picurl
	dim articleid
	dim classid
	dim errmsg
	dim founderr
	founderr=false
	if trim(request.form("txttitle"))="" then
  		founderr=true
  		errmsg="<li>名称不能为空</li>"
	end if
	if trim(request.form("txtcontent"))="" then
  		founderr=true
  		errmsg=errmsg+"<li>内容不能为空</li>"
	end if
	if trim(request.form("classid"))="" then
  		founderr=true
  		errmsg=errmsg+"<li>分类不能为空</li>"
	end if
	if founderr=false then
		title=trim(request.form("txttitle"))
		classid=request.form("classid")
        	picurl=trim(request.form("piccontent"))
		if request("htmlable")="yes" then
		content=request("txtcontent")
 		else
                content=replace(request.form("txtcontent"),"  ","&nbsp;&nbsp;")
		content=ubbcode(content)
		end if
		writefrom=request.form("writefrom")
		writer=request.form("writer")
		addtime=request.form("addtime")
	set rs=server.createobject("adodb.recordset")
	if request("action")="add" then
		call newsoft()
	elseif request("action")="edit" then
		call editsoft()
	else
		founderr=true
		errmsg=errmsg+"<li>没有选定参数</li>"
	end if
sub newsoft()
	sql="select * from article where (articleid is null)" 
	rs.open sql,conn,1,3
	rs.addnew
	rs("title")=title
	rs("content")=content
	rs("classid")=classid
    if picurl<>"" then
        rs("picurl")=picurl
    end if
    rs("writefrom")=writefrom
    rs("writer")=writer
   
	rs.update
	articleid=rs("articleid")
end sub
sub editsoft()
	sql="select * from article where articleid="&request("id")
	rs.open sql,conn,1,3
	rs("title")=title
	rs("content")=content
	rs("classid")=classid
    if picurl<>"" then
        rs("picurl")=picurl
    end if
    rs("writefrom")=writefrom
    rs("writer")=writer
    rs("addtime")=addtime
	rs.update
	articleid=rs("articleid")
end sub

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<html>
<head>
<title></title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body>
<div align="center"><center>
<br><br>
<table class="border" align=center width="50%" border="0" cellpadding="4" cellspacing="0" bordercolor="#999999">
  <tr align=center>
    <td width="100%" class="title"  height="20"><b>
<%if request("action")="add" then%>添加<%else%>修改<%end if%>资料成功</b></td>
  </tr>
  <tr>
    <td class="tdbg"><p align="left"><br>
        <b>
    资料</b>序号为：<%response.write "article"&articleid%><br>
        <b>
    资料</b>标题为：<%response.write title%></p>
    您可以进行其他操作
    </td>
  </tr>
</table>
</center></div>
<%response.write "<p align=center><a href='adminedit.asp?action=edit&classid="&request("classid")&"&classtype="&request("classtype")&"'>返回</a>"
else
	Error()
end if
%>
</body>
</html>
<%
sub Error()
		response.write "   <html><head><link rel='stylesheet' href='style.css'></head><body>"
		response.write "   <br><br><br>"
		response.write "    <table align='center' width='300' border='0' cellpadding='4' cellspacing='0' class='border'>"
		response.write "      <tr > "
		response.write "        <td class='title' colspan='2' height='15'> "
		response.write "          <div align='center'>由于以下的原因不能保存数据!</div>"
		response.write "        </td>"
		response.write "      </tr>"
		response.write "      <tr> "
		response.write "        <td align=center class='tdbg' colspan='2' height='23'> "
		response.write "          <br>"
		response.write errmsg& " <br><br>"
		response.write "        <a href='javascript:onclick=history.go(-1)'>返回</a>"        
		response.write "        <br><br></td>"
		response.write "      </tr>   </table></body></html>" 
end sub
%>