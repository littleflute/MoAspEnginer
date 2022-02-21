<%
	set cn=Server.CreateObject("ADODB.Connection")

	cn.open "DBQ="+server.mappath("btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"   

	if Request("A") <>"" then
		Account = Request("A")
		Password= Request("P")
		SQL = "Select * from 客户信息 Where 账号='" + Account + "' and 密码='" + Password + "'"
		
		set rs=cn.execute(SQL)
		
		if not rs.EOF then		 
			Response.Cookies("ContractWayAccount")=rs.fields("账号")
			Response.Cookies("ContractWayAccount").Expires="2010/1/1"		
			Response.Cookies("ContractWayTelephone")=rs.fields("电话")
			Response.Cookies("ContractWayTelephone").Expires="2010/1/1"		
			Response.Cookies("ContractWayEmails")=rs.fields("妹儿")
			Response.Cookies("ContractWayEmails").Expires="2010/1/1"
			Response.Cookies("ContractWayName")=rs.fields("姓名")
			Response.Cookies("ContractWayName").Expires="2010/1/1"
			Response.Cookies("ContractWayAddress")=rs.fields("住址")
			Response.Cookies("ContractWayAddress").Expires="2010/1/1"
			Response.Cookies("H19681019S19681231W19980412B69793682")="OK"
			Response.Cookies("H19681019S19681231W19980412B69793682").Expires="2010/1/1"
			Response.Redirect "money.asp"
		else
			ABC="你所输入的'密码'或'账号'有误"
		end if
		
		cn.close
	end if
	
	if Request("account") <> "" then
		SQL = "Select 账号 from 客户信息 where 账号='" & Request("account") & "'"
	
		set rs = cn.execute( SQL )
		if rs.EOF then
			Account = Request("account")
			Password= Request("key1")
		
			Mstr="Insert into 客户信息(账号,密码"
			Nstr="Values('" + Account + "','" + Password 
		
			if Request("myname") <> "" then
				Mstr=Mstr + ",姓名"
				Nstr=Nstr + + "','" + Request("myname")
			end if
		
			if Request("address") <> "" then
				Mstr=Mstr + ",住址"
				Nstr=Nstr + "','" + Request("address")
			end if

			if Request("email") <> "" then
				Mstr=Mstr + ",E-mail"
				Nstr=Nstr + "','" + Request("email")
			end if

			if Request("telephone") <> "" then
				Mstr=Mstr + ",电话"
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
			ABC="你所输的账号已有人注册"
		end if
		cn.close
	end if

%>


<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>用户验证</title>
</head>

<body>

<form method="POST" action>
  <p align="center"><%=ABC%></p>
  <p align="center"><input type="button" value="重新填写" name="B1" onClick= "history.go(-1)"><input type="button" value="关闭窗口" name="B2" onClick="self.close()"></p>
</form>

</body>
</html>