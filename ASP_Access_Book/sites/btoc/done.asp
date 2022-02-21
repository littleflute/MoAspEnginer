<%
	dim temp
	
	if request("OP") <> "修改" then
		'//新建
		M="insert into 商品信息(名称,作者,包装,分类,售价,定价,简介,说明,使用,标记,图标,时间) "
		N="values("
		
		N=N + "'" + request("T1") + "',"
		N=N + "'" + request("T2") + "',"
		N=N + "'" + request("T3") + "',"
				
		temp = split(request("D"),"/")
		
		N=N + "'" + temp(0) + "',"
		N=N + "'" + request("T5") + "',"
		N=N + "'" + request("T6") + "',"
			a=replace(request("S1")," ","&nbsp;")
			a=replace(a,vbCrlf,"<br>")
		N=N + "'" + a + "',"
			a=replace(request("S2")," ","&nbsp;")
			a=replace(a,vbCrlf,"<br>")
		N=N + "'" + a + "',"
			a=replace(request("S3")," ","&nbsp;")
			a=replace(a,vbCrlf,"<br>")
		N=N + "'" + a + "',"
		N=N + "'" + request("T7") + "',"
		N=N + "'" + temp(1)
		SQL="select count(*) as B from 商品信息 where 图标 like '" + temp(1) + "%'"

		set Conn=Server.CreateObject("ADODB.Connection")

		Conn.open "DBQ="+server.mappath("btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"  
		set rs = Conn.execute(SQL)
		
		if rs.fields("B") <> 0 then
			P= cstr(rs.fields("B")+1) + ".jpg"
			SQL = M + N + P +  + " '"
		else
			SQl = M + N + "1.jpg'"
			P ="1.jpg"
		end if
		
		SQL = SQL + ",'" + FormatDateTime(now) + "')"
		Conn.execute SQL
		AAA="请将你的图标命名为----" + temp(1) + P
		
		Conn.close

	else
		'//修改
		set Conn=Server.CreateObject("ADODB.Connection")

		Conn.open "DBQ="+server.mappath("btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"  

		SQL="update 商品信息 set 名称='" + request("T1") + "',"
		SQL=SQL + "作者='" + request("T2") + "',"
		SQL=SQL + "包装='" + request("T3") + "',"
		SQL=SQL + "分类='" + left(request("D"),4) + "',"
		SQL=SQL + "售价='" + request("T5") + "',"
		SQL=SQL + "定价='" + request("T6") + "',"
			a=replace(request("S1")," ","&nbsp;")
			a=replace(a,vbCrlf,"<br>")
		SQL=SQL + "简介='" + a + "',"
			a=replace(request("S2")," ","&nbsp;")
			a=replace(a,vbCrlf,"<br>")
		SQL=SQL + "说明='" + a + "',"
			a=replace(request("S3")," ","&nbsp;")
			a=replace(a,vbCrlf,"<br>")
		SQL=SQL + "使用='" + a + "',"
		SQL=SQL + "标记='" + request("T7") + "'"
		SQL=SQL + " where 名称='" + request("T1") + "'"
		Conn.execute(SQL)
		Conn.close
		AAA="修改已完成！"
		
	end if
%>


<htNl>

<head>
<Neta http-equiv="Content-Type" content="text/htNl; charset=gb2312">
<Neta naNe="GENERATOR" content="Nicrosoft FrontPage 4.0">
<Neta naNe="ProgId" content="FrontPage.Editor.DocuNent">
<title>新资料入库、旧资料管理</title>
</head>

<body>
<script language="javascript">
<!--
//-->
</script>
<form>
<table border="0" cellspacing="0" width="100%" cellpadding="0">
  <tr>
    <td width="100%" colspan="2">
      <p align="center"><%=AAA%></td>
  </tr>
  <tr>
    <td width="50%">
      <p align="right"><input type="button" value="回到主菜单" name="B1" onClick="self.location.replace('upld.asp')"></td>
    <td width="50%">
      <p align="left"><input type="button" value="向后退一步" name="B2" onClick="history.go(-1)"></td>
  </tr>
</table>
</form>
</body>    
    
