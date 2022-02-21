
<%
	FF = request("Flag")
	Num = request("Num")
	if Num = "" then
		SQL = "Select * from 订单管理 order by 订单号"
		set cn=Server.CreateObject("ADODB.Connection")

		cn.open "DBQ="+server.mappath("../btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 

		set rs = cn.Execute(SQL)
	else
		SQL = "Update 订单管理 set 标记='" + FF + "' where 订单号='" + Num + "'"
		set cn=Server.CreateObject("ADODB.Connection")

		cn.open "Provider=sqloledb;" & "Data Source=127.0.0.1;Initial Catalog=qiye_2;User Id=sa;Password=password;"  
		set rs = cn.Execute(SQL)
		SQL = "Select * from 订单管理 order by 订单号"
		set rs = cn.Execute(SQL)
	end if
		

%>


<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>洋洋购物网上购物订单管理系统</title>
</head>

<body topmargin="0" leftmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="74%">
    <tr>
      <td width="100%" colspan="5">
        <hr>
      </td>
    </tr>
    <tr>
      <td width="20%" align="center"><font face="方正姚体" size="3" color="#0000FF"><a href="newform.asp">未处理订单</a></font></td>
      <td width="20%" align="center"><font face="方正姚体" size="3" color="#0000FF"><a href="oldform.asp">已处理订单</a></font></td>
      <td width="20%" align="center"><font face="方正姚体" size="3" color="#0000FF"><a href="wasteform.asp">无效订单</a></font></td>
      <td width="20%" align="center"><font face="方正姚体" size="3" color="#0000FF"><a href="formsearch.asp">订单查询</a></font></td>
      <td width="20%" align="center"><font face="方正姚体" size="3" color="#0000FF"><a href="formmanage.asp">订单管理</a></font></td>
    </tr>
    <tr>
      <td width="100%" colspan="5">
        <hr><form name="Modi" action="formmanage.asp" method="POST"><input name="Flag" type="hidden"><input name="Num" type="hidden">
        <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#808080" height="39">
          <tr>
            <td width="3%" align="center" bgcolor="#FFFF66" height="19">&nbsp;</td>
            <td width="12%" align="center" bgcolor="#FFFF66" height="19"><font face="宋体" size="2"><b>订单号码</b></font></td>
            <td width="18%" align="center" bgcolor="#FFFF66" height="19"><font face="宋体" size="2"><b>用户名称</b></font></td>
            <td width="11%" align="center" bgcolor="#FFFF66" height="19"><font face="宋体" size="2"><b>配送方式</b></font></td>
            <td width="11%" align="center" bgcolor="#FFFF66" height="19"><font face="宋体" size="2"><b>付款方式</b></font></td>
            <td width="19%" align="center" bgcolor="#FFFF66" height="19"><font face="宋体" size="2"><b>填表时间</b></font></td>
            <td width="18%" align="center" bgcolor="#FFFF66" height="19"><font face="宋体" size="2"><b>完成时间</b></font></td>
            <td width="8%" align="center" bgcolor="#FFFF66" height="19"><font face="宋体" size="2"><b>标记</b></font></td>
          </tr>
<%
	Do while Not rs.EOF
	if rs.fields("标记") = "废" then
    	FF = "无"
    else
    	FF = "废"
    end if  			
%>
          <tr>
            <td width="3%" align="center" bgcolor="#FFFFFF" height="16"><font size="2" color="blue"><a style="cursor: hand" href onClick="
            	Modi.Flag.value='<%=FF%>';
            	Modi.Num.value='<%=rs.fields("订单号")%>';
            	Modi.submit()
            	">改</a></font></td>
            <td width="12%" align="center" bgcolor="#FFCCFF" height="16"><font face="宋体" size="2"><%=rs.fields("订单号")%></font></td>
            <td width="18%" align="center" bgcolor="#CCFFFF" height="16"><font face="宋体" size="2"><%=rs.fields("用户名")%></font></td>
            <td width="11%" align="center" bgcolor="#FFCCFF" height="16"><font face="宋体" size="2"><%=rs.fields("配送")%></font></td>
            <td width="11%" align="center" bgcolor="#CCFFFF" height="16"><font face="宋体" size="2"><%=rs.fields("付款")%></font></td>
            <td width="19%" align="center" bgcolor="#FFCCFF" height="16"><font face="宋体" size="2"><%=rs.fields("填写时间")%></font></td>
            <td width="18%" align="center" bgcolor="#CCFFFF" height="16"><font face="宋体" size="2"><%=rs.fields("完成时间")%></font></td>
            <td width="8%" align="center" bgcolor="#FFCCFF" height="16"><font face="宋体" size="2"><%=rs.fields("标记")%></font></td>
          </tr>
<%
	rs.Movenext
	loop
	cn.Close
%>
        </table></form>
        <p align="center">Copyright:北京公司</td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>

