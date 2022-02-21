<%
	Num = Request("T1")
	Nam = Request("T2")
	Time1 = Request("T3")
	Time2 = Request("T4")
	
	ff = 0 
	SQL = "Select * from 订单管理 where "
	if Num <> "" then 
		SQL = SQL + "订单号 like '%" + Num + "%'"
		ff = 1 
	end if
	
	if Nam <> "" then
		if ff = 1 then
			SQL = SQL + " and 用户名 like '%" + Nam + "%'"
		else	
			SQL = SQL + " 用户名 like '%" + Nam + "%'"
			ff = 1
		end if
	end if
	
	if Time1 <> "" then
		if ff = 1 then
			SQL = SQL + " and 填写时间 like '%" + Time1 + "%'"
		else
			SQL = SQL + " 填写时间 like '%" + Time1 + "%'"
			ff = 1
		end if
	end if
	
	if Time2 <> "" then
		if ff = 1 then
			SQL = SQL + " and 完成时间 like '%" + Time2 + "%'"
		else
			SQL = SQL + " 完成时间 like '%" + Time2 + "%'"
			ff = 1
		end if
	end if

	if ff = 1 then
		set cn=Server.CreateObject("ADODB.Connection")
	
		cn.open "DBQ="+server.mappath("../btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 

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
        <hr>
      </td>
    </tr>
    <tr>
      <td width="100%" colspan="5">
        <p align="center"><font face="方正姚体">订单查询系统</font></td>
    </tr>
    <tr>
      <td width="100%" colspan="5">
        <div align="center"><form name="SS" action="formsearch.asp">
          <table border="1" cellpadding="0" cellspacing="0" width="80%" bordercolorlight="#FFFFFF" bordercolordark="#808080" height="98">
            <tr>
              <td width="20%" bgcolor="#C0C0C0" align="center" height="21"><font size="2">订单号码</font></td>
              <td width="38%" height="21"><input type="text" name="T1" size="27"></td>
              <td width="42%" height="90" rowspan="4" valign="top" align="left"><font size="2">&nbsp;<span style="background-color: #FFFF00"><font color="#0000FF">使用方法:</font></span><br>
                &nbsp; 1、尽可能多地填写表单内容，以方便准确地找到所需内容。<br>
                &nbsp; 2、“订单号码”请使用大写字母。<br>
                &nbsp; 3、年月格式为：&nbsp; 01-08-18。<br> 
                &nbsp;&nbsp; 4、时间格式为：&nbsp; 13:20:50。<br> 
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                &nbsp;&nbsp; <input type="button" value="发送" name="B1" onClick="
                	if (SS.T1.value == '' && SS.T2.value == '' && SS.T3.value == '' && SS.T4.value == '') 
                	{
                		alert('请填写您需要查询的内容!')
                	} 
                	else
                	{
                		SS.submit()
                	}         
                
                	"> <input type="reset" value="重填" name="B2"></font></td> 
            </tr> 
            <tr> 
              <td width="20%" bgcolor="#C0C0C0" align="center" height="23"><font size="2">客户帐号</font></td> 
              <td width="38%" height="23"><input type="text" name="T2" size="27"></td> 
            </tr> 
            <tr> 
              <td width="20%" bgcolor="#C0C0C0" align="center" height="23"><font size="2">填写时间</font></td> 
              <td width="38%" height="23"><input type="text" name="T3" size="27"></td> 
            </tr> 
            <tr> 
              <td width="20%" bgcolor="#C0C0C0" align="center" height="23"><font size="2">完成时间</font></td> 
              <td width="38%" height="23"><input type="text" name="T4" size="27"></td> 
            </tr> 
          </table></form> 
        </div> 
  </center></td> 
    </tr> 
  <center> 
  <tr> 
      <td width="100%" colspan="5"> 
        <p align="center"><b><font face="楷体_GB2312"><br> 
        </font></b><font face="方正姚体">订单查询结果</font></p> 
      </td> 
  </tr> 
  <tr> 
      <td width="100%" colspan="5"> 
        <div align="center"> 
          <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#000080"> 
            <tr> 
              <td width="4%" bgcolor="#FF99FF">&nbsp;</td> 
              <td width="14%" bgcolor="#FF99FF" align="center"><font size="2">订单号码</font></td> 
              <td width="19%" bgcolor="#FF99FF" align="center"><font size="2">客户帐号</font></td> 
              <td width="12%" bgcolor="#FF99FF" align="center"><font size="2">配送方式</font></td> 
              <td width="12%" bgcolor="#FF99FF" align="center"><font size="2">付款方式</font></td> 
              <td width="20%" bgcolor="#FF99FF" align="center"><font size="2">填写时间</font></td> 
              <td width="19%" bgcolor="#FF99FF" align="center"><font size="2">完成时间</font></td> 
            </tr> 
<% 	
	if ff = 1 then
	Do while Not rs.EOF 
%> 
            <tr> 
              <td width="4%"><font size="2" color="blue"><a href onClick=" 
            	window.open('formtxt.asp?Num=<%=rs.fields("订单号")%>&Nam=<%=rs.fields("用户名")%>','formtxt','width=375,height=400,scrollbars=1')" style="CURSOR:hand">.&gt;&gt;</a> 
            	</font></td> 
              <td width="14%"><%=rs.fields("订单号")%></td> 
              <td width="19%"><%=rs.fields("用户名")%></td> 
              <td width="12%"><%=rs.fields("配送")%></td> 
              <td width="12%"><%=rs.fields("付款")%></td> 
              <td width="20%"><%=rs.fields("填写时间")%></td> 
              <td width="19%"><%=rs.fields("完成时间")%></td> 
            </tr> 
<% 
	rs.Movenext 
	Loop 
	cn.Close 
	end if
%> 
          </table> 
        </div> 
    </td> 
  </tr> 
  <tr> 
      <td width="100%" colspan="5"> 
    </td> 
  </tr> 
  </table> 
  </center> 
</div> 
 
</body> 
 
</html> 
