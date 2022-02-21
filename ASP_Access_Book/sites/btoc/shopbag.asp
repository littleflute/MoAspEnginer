
<%	 
  	dim Nstr(20)
  	  
  	if request("Clear") <> "" then
  		Response.Cookies("times")=0
  	end if
  	 
	if request("P") <> "" then 
  		Mstr=Request("P") 
  		Mstr=Mstr+"/"+Request("Price") +"/1" 
		T=Request.Cookies("times") 
		if T <> "" then 
			Times=Clng(T)+1 
		else 
			Times=1 
		end if 
		Response.Cookies("times")=cstr(Times) 
		Response.Cookies("MYShopBag"+ cstr(Times))=Mstr 
  	end if 
  	 
  	if request("D") <> "" then 
  		Position=Request("D") 
		Times=Clng(Request.Cookies("times")) 
		If Times=Position then 
			Response.Cookies("times")=0 
		Else 
			for i=Position to Times-1 
				Response.Cookies("MYShopBag"+cstr(i))=Request.Cookies("MYShopBag"+cstr(i+1)) 
			next 
			Response.Cookies("times")=Times-1 
		end if 
  	end if 
  	 	 
	C=request("C") 
	if C <> "" then
		Tstr=Request.Cookies("MYShopBag"+C)
		N=Request("ALLS"+C)
		Tstr=Left(Tstr,instrrev(Tstr,"/"))+N
		Response.Cookies("MYShopBag"+Request("C"))=Tstr
	end if
	
	
	Total=Request.Cookies("times") 
	if cstr(Total) = 0 then
		Total = 1
		Nstr(1) = "无/0/0/0"
	else
		for i = 1 to Clng(Total) 
			Nstr(i)=Request.Cookies("MYShopBag"+cstr(i)) 
		next	 
 	end if
%> 


<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>购物车中的商品</title>
<script src="java_function.js"></script>
</head>

<body >

 
<form name="MyShopbag"> <input type="hidden" Name="C"><input type="hidden" Name="D"><input type="hidden" Name="Clear">

<table width="514" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr valign="bottom"> 
    <td width="21"><img src="image/buy01.gif" width="21" height="59"></td>
      <td width="132" background="image/buyditu01.gif"><img src="image/buy02.gif" width="53" height="59"></td>
      <td align="right" background="image/buyditu01.gif" width="322"><img src="image/buy03.gif" width="206" height="35"></td>
      <td align="right" width="35"><img src="image/buy04.gif" width="35" height="35"></td>
  </tr>
  <tr> 
      <td width="21" background="image/buyditu04.gif" height="141">&nbsp;</td>
    <td colspan="2" valign="top" height="141" width="456">
    
<table border="1" width="100%" bordercolorlight="#FFFFFF" cellspacing="0" cellpadding="0" bordercolordark="#808080"> 
 
  <tr> 
    <td width="50%" bgcolor="#FFFF00" height="18"> 
      <p align="center"><font  size="2">商品名称</font></td> 
    <td width="14%" bgcolor="#FFFF00" height="18"> 
      <p align="center"><font  size="2">会员价</font></td> 
    <td width="5%" bgcolor="#FFFF00" height="18"> 
      <p align="center"><font  size="2">数量</font></td> 
    <td width="14%" bgcolor="#FFFF00" height="18"> 
      <p align="center"><font  size="2">小计</font></td> 
    <td width="17%" bgcolor="#FFFF00" height="18"> 
      <p align="center"><font  size="2">移到收藏夹</font></td> 
  </tr> 
 
 
  <%
  	dim S
  	dim CC
  	
  	for i = 1 to Total 
  		S=split(Nstr(i),"/")
  		CC=CC+Clng(S(1))*Clng(S(2))
  %> 
  <tr> 
    <td width="50%" align="center" height="18"><font  size="2"><%=S(0)%></font></td> 
    <td width="14%" align="center" height="18"><font  size="2"><%=S(1)%></font></td> 
    <td width="5%" align="center" height="18"> 
      <font  size="2"> 
      <input type="text" name="<%="ALLS"+cstr(i)%>" size="4" value="<%=S(2)%>" onChange="<%="MyShopbag.C.value=" + cstr(i) + "; MyShopbag.action='shopbag.asp'; MyShopbag.submit()"%>"> 
      </font> 
    </td> 
    <td width="14%" align="center" height="18"><font  size="2"><%=Clng(S(1))*Clng(S(2))%></font></td> 
    <td width="17%" align="center" height="18"><input type="button" name="<%=cstr(i)%>" value="删除" onClick="
    		if (confirm('是否真的删除此商品?'))
    		{
				MyShopbag.D.value=<%=cstr(i)%>;
				MyShopbag.action='shopbag.asp';
				MyShopbag.submit();		
			}">
  </tr> 
  <%
    next 
  %>
   
   
     
  <tr> 
    <td width="100%" colspan="5"> 
      　<table border="0" width="100%">
      		<tr>
      		<td width="60%"></td>
      		<td width="40%"><font  size="2" color="red">总计: <%=CC%>                       
              元</font></td>	                       
			</tr>   		                       
      	</table>                        
    </td>                        
  </tr>                        
                          
                          
                            
</table>                        
                       
<table  border="0" width="100%" bordercolorlight="#FFFFFF" cellspacing="0" cellpadding="0" bordercolordark="#808080" height="30">                       
	<tr>                       
		<td height="30">                       
          <p align="center"><input type="button" value="修改数量" name="modify"> <input type="button" name="clear" value="全部清除" onClick="MyShopbag.Clear.value='Y'; MyShopbag.action='shopbag.asp'; MyShopbag.submit()">                     
          <input type="button" name="go" value="去收银台" onClick="CheckCookie()"> <input type="button" value="继续购物" name="close" onClick="window.close()">                        
        </td>                       
	</tr>                       
	<tr>                       
		<td height="30">                       
              <p align="center">&nbsp;
            </td>                      
	</tr>                      
</table>                      
                      
                      
	</td>                      
      <td width="35" background="image/buyditu03.gif" height="141">&nbsp;</td>                      
  </tr>                      
  <tr valign="top">                       
      <td width="21"><img src="image/buy05.gif" width="21" height="21"></td>       
      <td colspan="2" background="image/buyditu02.gif" width="456">&nbsp;</td>                      
      <td width="35"><img src="image/buy06.gif" width="35" height="21"></td>                      
  </tr>                      
</table>                      
</form>                       
<table width="512" border="0" cellspacing="0" align="center">                    
  <tr>                    
    <td width="512" height="4" align="center">                    
      <hr width="100%" size="1" noshade>                    
    </td>                    
  </tr>                    
  <tr>                    
    <td width="512" height="4" align="center"><font size="2"><b>Copyright</b>:郑州亮中</font></td>                    
  </tr>                    
  <tr>                    
    <td width="512" height="4" align="center"><font size="2">　E-mail:<a href="mailto:xxxx@xxxx.com">技术支持</a></font></td>                   
  </tr>                   
</table>                   
</body>                     
                     
</html>                     
        
        
