<%
	set cn=Server.CreateObject("ADODB.Connection")

	cn.open "DBQ="+server.mappath("btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"   

	if Request("A") <>"" then
		Account = Request("A")
		Password= Request("P")
		SQL = "Select * from �ͻ���Ϣ Where �˺�='" + Account + "' and ����='" + Password + "'"
		
		set rs=cn.execute(SQL)
		
		if not rs.EOF then		 
			Response.Cookies("ContractWayAccount")=rs.fields("�˺�")
			Response.Cookies("ContractWayAccount").Expires="2010/1/1"		
			Response.Cookies("ContractWayTelephone")=rs.fields("�绰")
			Response.Cookies("ContractWayTelephone").Expires="2010/1/1"		
			Response.Cookies("ContractWayEmails")=rs.fields("�ö�")
			Response.Cookies("ContractWayEmails").Expires="2010/1/1"
			Response.Cookies("ContractWayName")=rs.fields("����")
			Response.Cookies("ContractWayName").Expires="2010/1/1"
			Response.Cookies("ContractWayAddress")=rs.fields("סַ")
			Response.Cookies("ContractWayAddress").Expires="2010/1/1"
			Response.Cookies("H19681019S19681231W19980412B69793682")="OK"
			Response.Cookies("H19681019S19681231W19980412B69793682").Expires="2010/1/1"
			Response.Redirect "money.asp"
		else
			ABC="���������'����'��'�˺�'����"
		end if
		
		cn.close
	end if
	
	if Request("account") <> "" then
		SQL = "Select �˺� from �ͻ���Ϣ where �˺�='" & Request("account") & "'"
	
		set rs = cn.execute( SQL )
		if rs.EOF then
			Account = Request("account")
			Password= Request("key1")
		
			Mstr="Insert into �ͻ���Ϣ(�˺�,����"
			Nstr="Values('" + Account + "','" + Password 
		
			if Request("myname") <> "" then
				Mstr=Mstr + ",����"
				Nstr=Nstr + + "','" + Request("myname")
			end if
		
			if Request("address") <> "" then
				Mstr=Mstr + ",סַ"
				Nstr=Nstr + "','" + Request("address")
			end if

			if Request("email") <> "" then
				Mstr=Mstr + ",E-mail"
				Nstr=Nstr + "','" + Request("email")
			end if

			if Request("telephone") <> "" then
				Mstr=Mstr + ",�绰"
				Nstr=Nstr + "','" + Request("telephone")
			end if

			SQL = Mstr + ") " + Nstr + "')"
				cn.execute SQL
				Response.Cookies("ContractWayAccount")=Request("Account")
				Response.Cookies("ContractWayAccount").Expires="2010/1/1"		
				Response.Cookies("ContractWayTelephone")=Request("telephone")
				Response.Cookies("ContractWayTelephone").Expires="2010/1/1"		
				Response.Cookies("ContractWayEmails")=Request("email")
				Response.Cookies("ContractWayEmails").Expires="2010/1/1"
				Response.Cookies("ContractWayName")=Request("myname")
				Response.Cookies("ContractWayName").Expires="2010/1/1"
				Response.Cookies("ContractWayAddress")=Request("address")
				Response.Cookies("ContractWayAddress").Expires="2010/1/1"
				Response.Cookies("H19681019S19681231W19980412B69793682")="OK"
				Response.Cookies("H19681019S19681231W19980412B69793682").Expires="2010/1/1"
				Response.Redirect "money.asp"
		else
			ABC="��������˺�������ע��"
		end if
		cn.close
	end if

%>


<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�û���֤</title>
</head>

<body>

<form method="POST" action>
  <p align="center"><%=ABC%></p>
  <p align="center"><input type="button" value="������д" name="B1" onClick= "history.go(-1)"><input type="button" value="�رմ���" name="B2" onClick="self.close()"></p>
</form>

</body>
</html>