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
		case 1 R = "�ʾֻ��"
		case 2 R = "���е��"
	end select
	TimeNow = FormatDateTime(now)
	
	'д�����ݿ�
	set cn=Server.CreateObject("ADODB.Connection")

	cn.open "DBQ="+server.mappath("btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"  
	
	SQL = "Insert into ��������(�û���,����,����,��дʱ��,���ʱ��,���) values('"
	SQL = SQL + Ac + "','"
	SQL = SQL + SS + "','"
	SQL = SQL + R + "','"
	SQL = SQL + TimeNow + "','��','��')"
	cn.Execute SQL
	
	SQL = "Select ID from �������� where �û���='" 
	SQL = SQL + Ac + "' and ��дʱ��='"
	SQL = SQL + TimeNow + "'"
	set rs = cn.Execute(SQL)
	
	Num="AB010" + CStr(rs.fields("ID"))
	SQL = "Update �������� set ������='"
	SQL = SQL + Num + "' where ID=" + cStr(rs.fields("ID"))
	'response.write SQL
	cn.Execute SQL	

  	dim S
  	dim C
  	for i = 1 to Clng(ts) 
  		S=split(PS(i),"/")
		SQL = "Insert into �������� values('"
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
<title>�ͻ���������ϵͳ</title>
<script language="javascript">
{
	alert("���Ĺ��ﶩ���ѷ���,лл����ʹ��!");
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
	      		<p align="left"><span style="background-color: #C0C0C0"><font  color="#0000FF" size="2">�ͻ���ϵ���ݣ�(�û���<%=Ac%>)</font></span>
	    	  </td>
	  		</tr>
		    <tr>
		      <td width="494">
	    	    <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#000000" bordercolor="#000000">
	        	  <tr>
	          	    <td width="10%" bgcolor="#FFFFFF"><font  size="2">������</font></td>
	            	<td width="12%" bgcolor="#FFFFFF"><p align="center"><font  size="2"><%=N%></font></td>
	          		<td width="9%" bgcolor="#FFFFFF"><font  size="2">�绰��</font></td>
	          		<td width="20%" bgcolor="#FFFFFF"><p align="center"><font  size="2"><%=T%></font></td>
	          		<td width="9%" bgcolor="#FFFFFF"><font  size="2">�ʼ���</font></td>
	          		<td width="40%" bgcolor="#FFFFFF"><p align="center"><font  size="2"><%=E%></font></td>
	        	  </tr>
	        	  <tr>
	          		<td width="10%" bgcolor="#FFFFFF"><font  size="2">סַ��</font></td>
	          		<td width="90%" colspan="5" bgcolor="#FFFFFF"><p align="center"><font  size="2"><%=A%></font></td>
	        	  </tr>
	      		</table>
	    	  </td>
	  		</tr>
	  		<tr>
	    	  <td width="494"><p align="left"> <font  color="#0000FF" size="2">�ͻ��������ݣ�</font></td>
	  		</tr>
	  		<tr>
	    	  <td width="494">
	      		<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#000000">
	        	  <tr>
	          		<td width="31%" align="center" bgcolor="#FFFFFF"><font  size="2">��Ʒ����</font></td>
	          		<td width="19%" align="center" bgcolor="#FFFFFF"><font  size="2">��Ա�۸�</font></td>
	          		<td width="25%" align="center" bgcolor="#FFFFFF"><font  size="2">ѡ������</font></td>
	          		<td width="25%" align="center" bgcolor="#FFFFFF"><font  size="2">�۸�С��</font></td>
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
	            <p align="right"><font  size="2">�۸��ܼƣ�<%=C%>Ԫ</font></td>
	        </tr>
	      </table>
	    </td>
	  </tr>
	  <tr>
	    <td width="494">
	    <font color="#0000FF"  size="2">��ѡ���ͷ�ʽ��</font>
	    </td>
	   </tr>
	  <tr>
	    <td width="494">
      
	      <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#000000">
	        <tr>
	          <td width="14%" bgcolor="#FFFFFF">
	            <p align="center"><font  size="2">���ͷ�ʽ��</font></td>
	          <td width="86%" bgcolor="#FFFFFF"><font size="2"><%=SS%></font></td>
	        </tr>
	      </table>
		    </td>
	   </tr>
	  <tr>
	    <td width="494">
	    <font  color="#0000FF" size="2">��ѡ���ʽ��</font>
	    </td>
	   </tr>
	  <tr>
	    <td width="494">
	    <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#000000">
	      <tr>
	        <td width="14%" bgcolor="#FFFFFF"><font  size="2">���ʽ��</font></td>
	        <td width="86%" bgcolor="#FFFFFF"><font size="2"><%=R%></font>
	      </tr>
	    </table>
	</tr>	
	<tr>
	<table>
	  <tr>
	    <td width="494" align="center">
        <p align="left"><font size="2" color="#FF0000">�û�ע������:<br>
        1.��ȷʵ�˶����Ķ������ݣ���������ϵ��ַ�������û��ʺš������������ݡ���<br>
        2.����������涩�������������壬�����ϰ����е绰��������ϵ��
        </font>
        <hr>
	    </td>
	  </tr>
	  <tr>
	    <td width="494" align="center">
	    <font size="2"><b>Copyright</b>:֣������</font>
	    </td>
	  </tr>
	  <tr>
	    <td width="494" align="center">
	    <font size="2">��Email:<a href="mailto:xxxx@xxxx.com">xxxx@xxxx.com</a></font>
	    </td>
	  </tr>
	</table>
	
</body>



