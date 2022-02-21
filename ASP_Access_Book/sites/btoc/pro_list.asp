<%
	set Conn=Server.CreateObject("ADODB.Connection")

	Conn.open "DBQ="+server.mappath("btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"    
	
	if request("D") <> "" then
		S=request("D")
		SS=split(S,"/")
		SQL = "delete from 商品信息 where 名称='" + SS(0) + "'and 图标='" + SS(1) + "'"
		Conn.execute SQL
	end if
	
	if request("C") <> "" then
		SQL="Select * from 商品信息 Where 分类='" + request("C") + "'"
	else
		SQL="Select * from 商品信息 Where 分类='战略策略'"  
	end if
%>



<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>

<body>


<form name="prolist" method="post"><input type="hidden" name="OP" value="修改"><div align="center">
<table border="0" width="770" cellspacing="1" cellpadding="0">
<tr>
  <td width="764" colspan="2">
  <font  size="3"><b>选择产品类别</b></font>
  </td>
</tr>
<tr>
  <td width="76">
  <font  size="2">听书馆</font>
  </td>
  <td width="688">
  <font  size="2" color="0000ff">&nbsp;
  	<a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=长篇评书'">长篇评书</a>&nbsp; 
  	<a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=传统相声'">传统相声</a>&nbsp; 
  	<a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=名著赏析'">名著赏析</a>&nbsp; 
  	<a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=其他内容'">其他内容</a>
  </font>   
  </td>   
</tr>   
<tr>   
  <td width="76">   
  <font size="2" >软件库</font>   
  </td>   
  <td width="688">
   <font size="2"  color="0000ff">&nbsp;
  	<a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=电脑学习'">电脑学习</a>&nbsp; 
  	<a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=外语学习'">外语学习</a>&nbsp; 
    <a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=家庭百科'">家庭百科</a>&nbsp; 
    <a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=多媒体类'">多媒体类</a>&nbsp; 
    <a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=查毒杀毒'">查毒杀毒</a>&nbsp; 
    <a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=新概念类'">新概念类</a>
   </font>   
  </td>   
</tr>   
<tr>   
  <td width="76">   
  <font size="2" >游戏厅</font>   
  </td>   
  <td width="688">
   <font size="2"  color="0000ff">&nbsp;
  	<a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=角色扮演'">角色扮演</a>&nbsp; 
  	<a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=动作冒险'">动作冒险</a>&nbsp; 
    <a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=战略策略'">战略策略</a>&nbsp; 
    <a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=模拟经营'">模拟经营</a>&nbsp; 
    <a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=益智休闲'">益智休闲</a>&nbsp; 
    <a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=体育竞技'">体育竞技</a>
   </font>   
  </td>   
</tr>   
</table>
<hr>
<table border="0" width="770" cellspacing="1" cellpadding="0">
<tr>   
  <td width="764">   
  </td>   
</tr>   
<tr>   
  <td width="602" align="center" valign="top">   
  <div align="center">   
    <center>   
    <table border="1" cellspacing="0" width="771" bordercolorlight="#FFFFFF" bordercolordark="#FFCCFF" cellpadding="0">   
      <tr>   
        <td width="52" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">图标</font></td>  
        <td width="110" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">名称</font></td>  
        <td width="50" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">作者</font></td>  
        <td width="54" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">分类</font></td>  
        <td width="54" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">包装</font></td>  
        <td width="30" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">原始定价</font></td>  
        <td width="28" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">网上售价</font></td>  
        <td width="287" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">说明</font></td>  
        <td width="43" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">标记</font></td>  
        <td width="40" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">操作</font></td>  
      </tr>  
  
<%  
	set rs=Conn.execute(SQL)  
	Do while not rs.EOF  
		Mstr="prolist.action='pro_list.asp?D=" + rs.fields("名称") + "/" + rs.fields("图标") + "'; prolist.submit()"  
		Lstr="sheet.asp?SQL=" + rs.fields("名称")
		Lstr=Lstr + "/" + rs.fields("图标")
		Lstr=Lstr + "/" + rs.fields("分类")
				
		Lstr="prolist.action='" + Lstr + "';  prolist.submit()"
%>  
      <tr>  
        <td width="52" valign="top" align="left" rowspan="2"><img src="pro-image/<%=rs.fields("图标")%>" width="50" height="60"></td>  
        <td width="110" valign="top" align="left" rowspan="2"><font  size="2"><%=rs.fields("名称")%></font></td>  
        <td width="50" valign="top" align="left" rowspan="2"><font  size="2"><%=rs.fields("作者")%></font></td>  
        <td width="54" valign="top" align="left" rowspan="2"><font  size="2"><%=rs.fields("分类")%></font></td>  
        <td width="54" valign="top" align="left" rowspan="2"><font  size="2"><%=rs.fields("包装")%></font></td>  
        <td width="30" valign="top" align="left" rowspan="2"><font  size="2"><%=rs.fields("定价")%></font></td>  
        <td width="28" valign="top" align="left" rowspan="2"><font  size="2"><%=rs.fields("售价")%></font></td>  
        <td width="287" valign="top" align="left" rowspan="2"><font  size="2"><%=rs.fields("说明")%></font></td>  
        <td width="43" valign="top" align="left" rowspan="2"><font  size="2"><%=rs.fields("标记")%></font></td>  
        <td width="40" valign="baseline" align="center"><input type="button" value="删除" name="B1" onClick="<%=Mstr%>"></td>  
      </tr>  
      <tr>  
        <td width="42" valign="baseline" align="center"><input type="button" value="修改" name="B2" onClick="<%=Lstr%>"></td>  
      </tr>  
  
        
<%  
	rs.movenext  
	loop  
	Conn.close  
%>  
  
  
  
    </table>  
    </center>  
  </div>  
  </td>  
</tr>  
<tr>  
  <td width="764">  
  </td>  
</tr>  
</table></div>  
</form>  
</body>  
  
</html>  
