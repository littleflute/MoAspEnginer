<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>

<body>
<div align="center"><form name="S" method="POST"><input type="hidden" name="OP" value="<%=request("OP")%>">

<%
	if request("OP") = "修改" then
		set Conn=Server.CreateObject("ADODB.Connection")

		Conn.open "DBQ="+server.mappath("btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"  
		
		response.write "修改纪录"
		S=request("SQL")
		T=split(S,"/")
		SQL="Select * from 商品信息 where 名称='" + T(0) + "' and 图标='" + T(1) + "' and 分类='" + T(2) + "'"		
		set rs=Conn.execute(SQL)
		
		TP=rs.fields("图标")
		T1=rs.fields("名称")
		T2=rs.fields("作者")
		T3=rs.fields("包装")
		T4=rs.fields("分类")
		T5=rs.fields("售价")
		T6=rs.fields("定价")
		T7=rs.fields("简介")
		T8=rs.fields("说明")
		T9=rs.fields("使用")
		C=rs.fields("标记")
		Conn.close
	else
		response.write "<font color=""blue""><u>新建记录</u></font>"
		TP="picture.jpg"
		C="最新"
	end if

%>

<table border="1" width="450" bordercolorlight="#FFFFFF" bordercolordark="#FFFF00">
  <tr>
    <td width="338"><font  size="2">名称: <input type="text" name="T1" size="41" value="<%=T1%>"></font></td>                                                              
    <td width="96" rowspan="4" align="right"><img src="pro-image/<%=TP%>" width="75" height="90">     
     
    </td>                                                              
  </tr>                                                              
  <tr>                                                              
    <td width="338"><font  size="2">作者: <input type="text" name="T2" size="41" value="<%=T2%>"></font></td>                                                              
  </tr>                                                              
  <tr>                                                              
    <td width="338"><font  size="2">包装: <input type="text" name="T3" size="17" value="<%=T3%>">&nbsp;&nbsp;&nbsp;&nbsp;    
      分类:                                                               
      <select size="1" name="D">                                                              
                      <option selected value="<%=T4%>"><%=T4%></option>                                                    
                      <option value="长篇评书/C">长篇评书</option>
                      <option value="传统相声/C">传统相声</option>
                      <option value="名著赏析/C">名著赏析</option>
                      <option value="其他内容/C">其他内容</option>
                      <option value="电脑学习/S">电脑学习</option>
                      <option value="外语学习/S">外语学习</option>
                      <option value="家庭百科/S">家庭百科</option>
                      <option value="多媒体类/S">多媒体类</option>
                      <option value="查毒杀毒/S">查毒杀毒</option>
                      <option value="新概念类/S">新概念类</option>
                      <option value="角色扮演/G">角色扮演</option>
                      <option value="动作冒险/G">动作冒险</option>
                      <option value="战略策略/G">战略策略</option>
                      <option value="模拟经营/G">模拟经营</option>
                      <option value="益智休闲/G">益智休闲</option>
                      <option value="体育竞技/G">体育竞技</option>
                    </select></font></td>
  </tr>
  <tr>
    <td width="338"><font  size="2">网上售价: <input type="text" name="T5" size="17" value="<%=T5%>">&nbsp;&nbsp;<br>
      原始定价:                                                               
      <input type="text" name="T6" size="11" value="<%=T6%>"></font></td>                                                              
  </tr>                                                              
  <tr>                                                            
    <td width="440" valign="top" align="left" colspan="2"><font  size="2">简介</font><font  size="2">:</font></td>                                                             
  </tr>                                                           
  <tr>                                                           
    <td width="440" valign="top" align="left" colspan="2"><textarea rows="4" name="S1" cols="60"><%=T7%></textarea></td>                                                             
  </tr>                                                           
  <tr>                                                             
    <td width="440" valign="top" align="left" colspan="2"><font  size="2">详细说明:</font></td>                                                             
  </tr>                                                           
  <tr>                                                             
    <td width="440" valign="top" align="left" colspan="2"><textarea rows="4" name="S2" cols="60"><%=T8%></textarea></td>                                                             
  </tr>                                                           
  <tr>                                                             
    <td width="440" valign="top" align="left" colspan="2"><font  size="2">使用说明：</font></td>                                                             
  </tr>                                                           
  <tr>                                                             
    <td width="440" valign="top" align="left" colspan="2"><textarea rows="4" name="S3" cols="60"><%=T9%></textarea></td>                                                             
  </tr>                                                           
  <tr>                                                             
    <td width="338" valign="top" align="left"><font  size="2">标记:                                                               
      <input type="text" name="T7" size="18" value="<%=C%>"></font></td>                                                              
    <td width="96" valign="top" align="right"><input type="button" value="确定" name="B1" onClick="                                
			if (S.T1.value=='' || S.T2.value=='' || S.T3.value=='' || S.D.value=='' || S.T5.value=='' || S.T6.value=='' || S.T7.value=='' || S.S1.value=='' || S.S2.value=='' || S.S3.value=='')                                 
    			{                                 
    				alert('请填写所有的空白！')                                 
   				}                               
   			else                         
   				{                         
   					S.action='done.asp'                         
   					S.submit()                         
   				}">                                                          
      <input type="reset" value="重填" name="B2"></td>                                                            
  </tr>                                                            
</table>                                                          
                                                          
</form>                                                            
</div>                                                            
</body>                                                            
                                                            
</html>                                                            
