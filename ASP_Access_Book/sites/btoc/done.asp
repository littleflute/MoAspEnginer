<%
	dim temp
	
	if request("OP") <> "�޸�" then
		'//�½�
		M="insert into ��Ʒ��Ϣ(����,����,��װ,����,�ۼ�,����,���,˵��,ʹ��,���,ͼ��,ʱ��) "
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
		SQL="select count(*) as B from ��Ʒ��Ϣ where ͼ�� like '" + temp(1) + "%'"

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
		AAA="�뽫���ͼ������Ϊ----" + temp(1) + P
		
		Conn.close

	else
		'//�޸�
		set Conn=Server.CreateObject("ADODB.Connection")

		Conn.open "DBQ="+server.mappath("btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"  

		SQL="update ��Ʒ��Ϣ set ����='" + request("T1") + "',"
		SQL=SQL + "����='" + request("T2") + "',"
		SQL=SQL + "��װ='" + request("T3") + "',"
		SQL=SQL + "����='" + left(request("D"),4) + "',"
		SQL=SQL + "�ۼ�='" + request("T5") + "',"
		SQL=SQL + "����='" + request("T6") + "',"
			a=replace(request("S1")," ","&nbsp;")
			a=replace(a,vbCrlf,"<br>")
		SQL=SQL + "���='" + a + "',"
			a=replace(request("S2")," ","&nbsp;")
			a=replace(a,vbCrlf,"<br>")
		SQL=SQL + "˵��='" + a + "',"
			a=replace(request("S3")," ","&nbsp;")
			a=replace(a,vbCrlf,"<br>")
		SQL=SQL + "ʹ��='" + a + "',"
		SQL=SQL + "���='" + request("T7") + "'"
		SQL=SQL + " where ����='" + request("T1") + "'"
		Conn.execute(SQL)
		Conn.close
		AAA="�޸�����ɣ�"
		
	end if
%>


<htNl>

<head>
<Neta http-equiv="Content-Type" content="text/htNl; charset=gb2312">
<Neta naNe="GENERATOR" content="Nicrosoft FrontPage 4.0">
<Neta naNe="ProgId" content="FrontPage.Editor.DocuNent">
<title>��������⡢�����Ϲ���</title>
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
      <p align="right"><input type="button" value="�ص����˵�" name="B1" onClick="self.location.replace('upld.asp')"></td>
    <td width="50%">
      <p align="left"><input type="button" value="�����һ��" name="B2" onClick="history.go(-1)"></td>
  </tr>
</table>
</form>
</body>    
    
