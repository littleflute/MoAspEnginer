<%
	pro = request("P")
	pic = request("T")
	
	SQL = "select * from 商品信息 where 名称='" + pro + "' and 图标='" + pic + "'"
	SET Conn=Server.CreateObject("ADODB.Connection") 

	Conn.open "DBQ="+server.mappath("btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"     
	set rs=Conn.execute(SQL) 
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>商品的详细资料</title>
</head>

<body><div><center>
<table border="1" bordercolor="#FFFF00"  bordercolorlight="#FFFFFF" bordercolordark="#0000FF"><tr><td>

<table border="0" cellpadding="0" cellspacing="0" width="460" height="87">
  <tr>
    <td width="481" height="10" bgcolor="#D8D8D8"><b><font size="2">商品名称：</font></b><font color="#cc000" size="2"><b><%=rs.fields("名称")%></b></font></td>
    <td width="67" height="32" rowspan="3"><img border="0" src="pro-image/<%=rs.fields("图标")%>" width="70" height="90"></td>
  </tr>
  <tr>
    <td width="481" height="4"><font size="2"><b>所属类别：</b><%=rs.fields("分类")%></font></td>
  </tr>
  <tr>
    <td width="481" height="13" bgcolor="#D8D8D8"><font size="2"><b>商品价格：</b>定价 <strike><font color="#CC0000"><%=rs.fields("定价")%>元</font></strike>&nbsp;&nbsp;&nbsp;&nbsp;     
      售价 <font color="#008000"><%=rs.fields("售价")%>元</font></font></td>    
  </tr>    
  <tr>    
    <td width="587" height="1" colspan="2"></td>    
  </tr>    
  <tr>   
    <td width="587" height="47" colspan="2" valign="top" bgcolor="#FF8040"><b><font size="2">内容介绍：<br>   
      </font></b><font size="2"><%=rs.fields("说明")%></font></td>    
  </tr>   
  <tr>    
    <td width="587" height="13" colspan="2"></td>    
  </tr>    
  <tr>    
    <td width="587" height="42" colspan="2" valign="top" bgcolor="#D8D8D8"><b><font size="2">使用说明：</font></b><font size="2"><br>    
      <%=rs.fields("使用")%></font></td>   
  </tr>   
  <tr>   
    <td width="587" height="10" colspan="2">   
      <p align="right"><font size="2" color="#0000FF"><u><a href style="CURSOR: hand" onClick="window.close()">返回主窗口</a></u></font></td>   
  </tr>   
</table>   
   
</td></tr></table>   
<%  
	Conn.close  
%>   
</center></div>  
</body>   
   
</html>   
