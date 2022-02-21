<%
	Ac= request.cookies("ContractWayAccount")
	N = request.cookies("ContractWayName")
	T = request.cookies("ContractWayTelephone")
	E = request.cookies("ContractWayEmails")
	A = request.cookies("ContractWayAddress")

	dim PS(20)
	
	ts = request("times")
	response.cookies("times")=0
	for i = 1 to Clng(ts)
		PS(i) = request("MYShopBag"+cstr(i))
	next 
	
	SS = request("Sendit")
	R = request("Recieve")
	select case R
		case 1 R = "邮局汇款"
		case 2 R = "银行电汇"
	end select
	TimeNow = FormatDateTime(now)
	
	'写入数据库
	set cn=Server.CreateObject("ADODB.Connection")

	cn.open "DBQ="+server.mappath("btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"  
	
	SQL = "Insert into 订单管理(用户名,配送,付款,填写时间,完成时间,标记) values('"
	SQL = SQL + Ac + "','"
	SQL = SQL + SS + "','"
	SQL = SQL + R + "','"
	SQL = SQL + TimeNow + "','无','无')"
	cn.Execute SQL
	
	SQL = "Select ID from 订单管理 where 用户名='" 
	SQL = SQL + Ac + "' and 填写时间='"
	SQL = SQL + TimeNow + "'"
	set rs = cn.Execute(SQL)
	
	Num="AB010" + CStr(rs.fields("ID"))
	SQL = "Update 订单管理 set 订单号='"
	SQL = SQL + Num + "' where ID=" + cStr(rs.fields("ID"))
	'response.write SQL
	cn.Execute SQL	

  	dim S
  	dim C
  	for i = 1 to Clng(ts) 
  		S=split(PS(i),"/")
		SQL = "Insert into 订单内容 values('"
		SQL = SQL + Num + "','"
		SQL = SQL + S(0) + "','"
		SQL = SQL + S(2) + "','"
		SQL = SQL + S(1) + "')"
		cn.Execute SQL
	next
	cn.Close
%>

<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>客户订单发送系统</title>
<script language="javascript">
{
	alert("您的购物订单已发出,谢谢您的使用!");
}
</script>
</head>

<body>
  <div align="center">
	<table bordercolorlight="#FFFFFF" border="1" bordercolordark="#000000" height="210" cellspacing="0" cellpadding="0">
	  <tr>
	    <td><img border="0" src="indeximage/banner.jpg" width="495" height="55">
	    </td>
	  </tr>
	  <tr>
	    <td height="204" valign="top">
	      <table border="0" cellpadding="0" cellspacing="0" width="496" bgcolor="#C0C0C0">
	   	    <tr>
	  	      <td width="494" bgcolor="#C0C0C0">
	      		<p align="left"><span style="background-color: #C0C0C0"><font  color="#0000FF" size="2">客户联系内容：(用户名<%=Ac%>)</font></span>
	    	  </td>
	  		</tr>
		    <tr>
		      <td width="494">
	    	    <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#000000" bordercolor="#000000">
	        	  <tr>
	          	    <td width="10%" bgcolor="#FFFFFF"><font  size="2">姓名：</font></td>
	            	<td width="12%" bgcolor="#FFFFFF"><p align="center"><font  size="2"><%=N%></font></td>
	          		<td width="9%" bgcolor="#FFFFFF"><font  size="2">电话：</font></td>
	          		<td width="20%" bgcolor="#FFFFFF"><p align="center"><font  size="2"><%=T%></font></td>
	          		<td width="9%" bgcolor="#FFFFFF"><font  size="2">邮件：</font></td>
	          		<td width="40%" bgcolor="#FFFFFF"><p align="center"><font  size="2"><%=E%></font></td>
	        	  </tr>
	        	  <tr>
	          		<td width="10%" bgcolor="#FFFFFF"><font  size="2">住址：</font></td>
	          		<td width="90%" colspan="5" bgcolor="#FFFFFF"><p align="center"><font  size="2"><%=A%></font></td>
	        	  </tr>
	      		</table>
	    	  </td>
	  		</tr>
	  		<tr>
	    	  <td width="494"><p align="left"> <font  color="#0000FF" size="2">客户订单内容：</font></td>
	  		</tr>
	  		<tr>
	    	  <td width="494">
	      		<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#000000">
	        	  <tr>
	          		<td width="31%" align="center" bgcolor="#FFFFFF"><font  size="2">商品名称</font></td>
	          		<td width="19%" align="center" bgcolor="#FFFFFF"><font  size="2">会员价格</font></td>
	          		<td width="25%" align="center" bgcolor="#FFFFFF"><font  size="2">选购数量</font></td>
	          		<td width="25%" align="center" bgcolor="#FFFFFF"><font  size="2">价格小计</font></td>
	        	  </tr>
<%
  	for i = 1 to Clng(ts) 
  		S=split(PS(i),"/")
  		C=C+Clng(S(1)) * Clng(S(2))
%>
	        <tr>
	          <td width="31%" align="center" bgcolor="#FFFFFF"><font  size="2"> <%= S(0)%></font></td>
	          <td width="19%" align="center" bgcolor="#FFFFFF"><font  size="2"> <%= S(1)%></font></td>
	          <td width="25%" align="center" bgcolor="#FFFFFF"><font  size="2"> <%= S(2)%></font></td>
	          <td width="25%" align="center" bgcolor="#FFFFFF"><font  size="2"> <%= cstr(clng(S(1))*clng(S(2)))%></font></td>
	        </tr>
        
<%    next %>
        
	        <tr>
	          <td width="100%" colspan="4" bgcolor="#FFFFFF">
	            <p align="right"><font  size="2">价格总计：<%=C%>元</font></td>
	        </tr>
	      </table>
	    </td>
	  </tr>
	  <tr>
	    <td width="494">
	    <font color="#0000FF"  size="2">所选配送方式：</font>
	    </td>
	   </tr>
	  <tr>
	    <td width="494">
      
	      <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#000000">
	        <tr>
	          <td width="14%" bgcolor="#FFFFFF">
	            <p align="center"><font  size="2">配送方式：</font></td>
	          <td width="86%" bgcolor="#FFFFFF"><font size="2"><%=SS%></font></td>
	        </tr>
	      </table>
		    </td>
	   </tr>
	  <tr>
	    <td width="494">
	    <font  color="#0000FF" size="2">所选付款方式：</font>
	    </td>
	   </tr>
	  <tr>
	    <td width="494">
	    <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#000000">
	      <tr>
	        <td width="14%" bgcolor="#FFFFFF"><font  size="2">付款方式：</font></td>
	        <td width="86%" bgcolor="#FFFFFF"><font size="2"><%=R%></font>
	      </tr>
	    </table>
	</tr>	
	<tr>
	<table>
	  <tr>
	    <td width="494" align="center">
        <p align="left"><font size="2" color="#FF0000">用户注意事项:<br>
        1.请确实核对您的订单内容，包括“联系地址”、“用户帐号”、“订单内容”等<br>
        2.如果您对上面订单的内容有疑义，请马上按下列电话与我们联系。
        </font>
        <hr>
	    </td>
	  </tr>
	  <tr>
	    <td width="494" align="center">
	    <font size="2"><b>Copyright</b>:郑州亮中</font>
	    </td>
	  </tr>
	  <tr>
	    <td width="494" align="center">
	    <font size="2">　Email:<a href="mailto:xxxx@xxxx.com">xxxx@xxxx.com</a></font>
	    </td>
	  </tr>
	</table>
	
</body>



