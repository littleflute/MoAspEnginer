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
	dim filename,filename1,filename2
	dim showname
	dim bb
	dim club
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
	dim hots
	dim hide
	dim founderr
	founerr=false
	if trim(request.form("txtshowname"))="" then
  		founderr=true
  		errmsg=errmsg+"<li>ӰƬ���Ʋ���Ϊ��</li>"
	end if
	if trim(request.form("txtnote"))="" then
  		founderr=true
  		errmsg=errmsg+"<li>ӰƬ��鲻��Ϊ��</li>"
	end if

	if founderr=false then
	    movie=request("movie")
		if request.form("club")="on" then
			club=1
		else
			club=0
		end if
		filename=request("txtfilename")
		filename1=request("txtfilename1")
		filename2=request("txtfilename2")
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
	rs("club")=club
	rs("filename")=filename
	if filename1<>"" then
	rs("filename1")=filename1
	end if
	if filename2<>"" then
	rs("filename2")=filename2
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
	rs("lasthits")=date()
	rs("dateandtime")=Now()
	rs("dayhits")=0
	rs("weekhits")=0
	rs("hits")=0
	rs.update
end sub
sub editsoft()
	sql="select * from download where id="&request("id")
	rs.open sql,conn,1,3
	rs("movie")=movie
	rs("club")=club
	rs("filename")=filename
	rs("filename1")=filename1
	rs("filename2")=filename2
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
	rs("lasthits")=date()
	rs("dateandtime")=Now()
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

<body bgcolor="#468ea3">
<br><br>
  
<table width="50%" border="0" cellpadding="3" cellspacing="1" align="center">
  <tr>
    
<td width="100%" height="20" bgcolor="#145f74">
<p align="center"><font color="#FFFFFF">
 <%if request("action")="add" then%>
 ���
 <%else%>
 �޸�
 <%end if%>
 ����ɹ�</font>
</td>
  </tr>
  <tr>
    
<td width="100%" bgcolor="#a5d0dc">
<p align="left"><br>
    ��ʾ����Ϊ��<%=request.form("txtshowname")%>&nbsp;<%=bb%><br>
    ���ص�ַ1Ϊ��<%=filename%><br>
    ���ص�ַ2Ϊ��<%=filename1%><br>
    ���ص�ַ3Ϊ��<%=filename2%><br>
    </p>
    �����Խ�����������
    </td>
  </tr>
</table>
<%
else
 response.write "�������µ�ԭ���ܱ������ݣ�"
 response.write errmsg
end if
%>

</body>
</html>
