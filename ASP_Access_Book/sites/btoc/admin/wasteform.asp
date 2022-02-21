<%
	set cn=Server.CreateObject("ADODB.Connection")

	cn.open "DBQ="+server.mappath("../btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"   

	SQL = "Select * from 订单管理 where 标记='废' order by 订单号"
	set rs = cn.Execute(SQL)

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
        <hr>
        <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#808080">
          <tr>
            <td width="5%" align="center" bgcolor="#FFFF66">&nbsp;</td>
            <td width="19%" align="center" bgcolor="#FFFF66"><font face="宋体" size="2"><b>订单号码</b></font></td>
            <td width="22%" align="center" bgcolor="#FFFF66"><font face="宋体" size="2"><b>用户名称</b></font></td>
            <td width="17%" align="center" bgcolor="#FFFF66"><font face="宋体" size="2"><b>配送方式</b></font></td>
            <td width="17%" align="center" bgcolor="#FFFF66"><font face="宋体" size="2"><b>付款方式</b></font></td>
            <td width="20%" align="center" bgcolor="#FFFF66"><font face="宋体" size="2"><b>填写时间</b></font></td>
          </tr>
<%
	Do while Not rs.EOF
%>
          <tr>
            <td width="5%" align="center" bgcolor="#FFFFFF"><font size="2" color="blue"><a href onClick="
            	window.open('formtxs.asp?Num=<%=rs.fields("订单号")%>&Nam=<%=rs.fields("用户名")%>','formtxt','width=375,height=400,scrollbars=1')" style="CURSOR:hand">.&gt;&gt;</a>
            	</font></td>
            <td width="19%" align="center" bgcolor="#FFCCFF"><font face="宋体" size="2"><%=rs.fields("订单号")%></font></td>
            <td width="22%" align="center" bgcolor="#CCFFFF"><font face="宋体" size="2"><%=rs.fields("用户名")%></font></td>
            <td width="17%" align="center" bgcolor="#FFCCFF"><font face="宋体" size="2"><%=rs.fields("配送")%></font></td>
            <td width="17%" align="center" bgcolor="#CCFFFF"><font face="宋体" size="2"><%=rs.fields("付款")%></font></td>
            <td width="20%" align="center" bgcolor="#FFCCFF"><font face="宋体" size="2"><%=rs.fields("填写时间")%></font></td>
          </tr>
<%
	rs.Movenext
	loop
	cn.Close
%>
        </table>
        <p align="center">Copyright:北京公司</td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>
