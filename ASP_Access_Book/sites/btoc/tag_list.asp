<%
	set Conn=Server.CreateObject("ADODB.Connection")

	Conn.open "DBQ="+server.mappath("btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"    
	
	if request("D") <> "" then
		S=request("D")
		SS=split(S,"/")
		SQL = "Update 商品信息 set 标记='无' where 名称='" + SS(0) + "'and 图标='" + SS(1) + "'"
		Conn.execute SQL
	end if
	
	if request("C") <> "" then
		S=request("C")
		SS=split(S,"/")
		SQL = "Update 商品信息 set 标记='" + SS(2) + "' where 名称='" + SS(0) + "'and 图标='" + SS(1) + "'"
		Conn.execute SQL
	end if
%>



<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>

<body>


<form name="taglist" method="post"><div align="center">
<table border="0" width="592" cellspacing="1" cellpadding="0">
<tr>
  <td width="586">
  <font face="隶书" size="4">
  首页产品标识（标记）管理</font>
  </td>
</tr>
<tr>
  <td width="586">
  <div align="center">
    <center>
    <table border="1" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#FFCCFF" cellpadding="0">
      <tr>
        <td width="13%" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">产品图标</font></td>
        <td width="39%" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">产品名称</font></td>
        <td width="14%" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">产品分类</font></td>
        <td width="14%" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">产品标识</font></td>
        <td width="10%" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">操作1</font></td>
        <td width="10%" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">操作2</font></td>
      </tr>

<%
	SQL="Select * from 商品信息 Where 标记<>'无'"
	set rs=Conn.execute(SQL)
	Do while not rs.EOF
		Mstr="taglist.action='tag_list.asp?D=" + rs.fields("名称") + "/" + rs.fields("图标") + "'; taglist.submit()"
		Lstr="var answer=prompt('请输入标记： 最新、推荐、或清仓','最新|推荐|清仓'); taglist.action='tag_list.asp?C=" + rs.fields("名称") + "/" + rs.fields("图标") + "/' + answer; taglist.submit()"                             
%>
      <tr>
        <td width="13%" valign="baseline" align="center"><img src="pro-image/<%=rs.fields("图标")%>" width="50" height="60"></td>
        <td width="39%" valign="baseline" align="left"><font  size="2"><%=rs.fields("名称")%></font></td>
        <td width="14%" valign="baseline" align="center"><font  size="2"><%=rs.fields("分类")%></font></td>
        <td width="14%" valign="baseline" align="center"><font  size="2"><%=rs.fields("标记")%></font></td>
        <td width="10%" valign="baseline" align="center"><input type="button" value="删除" name="B1" onClick="<%=Mstr%>"></td>
        <td width="10%" valign="baseline" align="center"><input type="button" value="修改" name="B2" onClick="<%=Lstr%>"></td>
      </tr>
      
<%
	rs.movenext
	loop
	Conn.close
%>


    </table>
    </center>
  </div>
  </td>
</tr>
<tr>
  <td width="586">
  <p align="right"><input type="button" value="返回主菜单" name="B3" onClick="location.replace('upld.asp')">
  </td>
</tr>
</table></div>
</form>
</body>

</html>

