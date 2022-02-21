
<%	 
  	dim Nstr(20)
  	  
  	if request("Clear") <> "" then
  		Response.Cookies("f_times")=0
  		Response.Cookies("f_times").Expires="2010/1/1"
  	end if
  	 
	if request("P") <> "" then 
  		Mstr=Request("P") 
  		Mstr=Mstr+"/"+Request("Price") +"/1/" + Request("T") 
		T=Request.Cookies("f_times") 
		if T <> "" then 
			f_times=Clng(T)+1 
		else 
			f_times=1 
		end if 
		Response.Cookies("f_times")=cstr(f_times) 
  		Response.Cookies("f_times").Expires="2010/1/1"
		Response.Cookies("MYFavoriteBag"+ cstr(f_times))=Mstr 
		Response.Cookies("MYFavoriteBag"+ cstr(f_times)).Expires="2010/1/1" 
  	end if 
  	 
  	if request("D") <> "" then 
  		Position=Request("D") 
		f_times=Clng(Request.Cookies("f_times")) 
		If f_times=Position then 
			Response.Cookies("f_times")=0 
	  		Response.Cookies("f_times").Expires="2010/1/1"
		Else 
			for i=Position to f_times-1 
				Response.Cookies("MYFavoriteBag"+cstr(i))=Request.Cookies("MYFavoriteBag"+cstr(i+1)) 
				Response.Cookies("MYFavoriteBag"+cstr(i)).Expires="2010/1/1"
			next 
			Response.Cookies("f_times")=f_times-1 
 	 		Response.Cookies("f_times").Expires="2010/1/1"
		end if 
  	end if 
  	 
	if request("M") <> "" then
		P=request("M")
		T=Clng(Request.Cookies("f_times"))
		T0=Request.Cookies("times")
		if T0 <> "" then
			Response.Cookies("times")=Clng(T0) + 1			
			Response.Cookies("MyShopbag" + Request.Cookies("times")) = Request.Cookies("MYFavoriteBag" + P)
		else
			Response.Cookies("times")=1
			Response.Cookies("MyShopbag1") = Request.Cookies("MYFavoriteBag" + P)
		end if
						
		if T = P then
			Response.Cookies("f_times")=0
			Response.Cookies("f_times").Expires="2010/1/1"
		else
			for i = P to T-1
				Response.Cookies("MYFavoriteBag"+cstr(i))=Request.Cookies("MYFavoriteBag"+cstr(i+1)) 
				Response.Cookies("MYFavoriteBag"+cstr(i)).Expires="2010/1/1"
			next 
			Response.Cookies("f_times")=T-1 
  			Response.Cookies("f_times").Expires="2010/1/1"
		end if
	end if
	
	
	Total=Request.Cookies("f_times") 
	if Total = "" then 
		Total = 1 
		Nstr(1)="无/无/无/picture.jpg"
	else	
		for i = 1 to Clng(Total) 
			Nstr(i)=Request.Cookies("MYFavoriteBag"+cstr(i)) 
		next	 
 	end if
%> 


<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>收藏夹中的商品</title>
<script language="javascript">
<!--
	function DeleteIt(B)
	{
		if (confirm("是否真的删除此商品？"))
		{
			Mybag.D.value=B
			Mybag.action="favorite.asp"
			Mybag.submit()		
		}
	}
	
	function MoveToShop(A)
	{
		Mybag.M.value=A
		Mybag.action="favorite.asp"
		Mybag.submit()
	}
//-->
</script>
</head>

<body >

 
<form name="Mybag"><input type="hidden" Name="D"><input type="hidden" Name="Clear"><input type="hidden" Name="M">

<table width="514" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr valign="bottom"> 
      <td width="21"><img src="image/favorite01.gif" width="21" height="59"></td>
      <td width="132" background="image/buyditu01.gif"><img src="image/favorite02.gif" width="53" height="59"></td>
      <td align="right" background="image/buyditu01.gif" width="322"><img src="image/favorite03.gif" width="206" height="35"></td>
      <td align="right" width="35"><img src="image/buy04.gif" width="35" height="35"></td>
  </tr>
  <tr> 
      <td width="21" background="image/buyditu04.gif" height="141">&nbsp;</td>
    <td colspan="2" valign="top" height="141" width="456">
    
        <table border="1" width="100%" bordercolorlight="#FFFFFF" cellspacing="0" cellpadding="0" bordercolordark="#808080">
          <tr bgcolor="#FFFF99"> 
            <td width="14%" height="18"> 
              <p align="center"><font size="2" >封面</font></td> 
            <td width="36%" height="18"> 
              <p align="center"><font  size="2">商品名称</font></td> 
            <td width="14%" height="18"> 
              <p align="center"><font  size="2">会员价</font></td> 
            <td width="31%" height="18"> 
              <p align="center"><font  size="2">收藏夹操作</font></td> 
  </tr> 
 
 
  <%
  	dim S
  	dim CC
  	
  	for i = 1 to Total 
  		S=split(Nstr(i),"/")
  %> 
  <tr> 
    <td width="14%" align="center" height="18"><font  size="2"><img src="pro-image/<%=S(3)%>" width="40" height="50"></font></td> 
    <td width="36%" align="center" height="18"><font  size="2"><%=S(0)%></font></td> 
    <td width="14%" align="center" height="18"><font  size="2"><%=S(1)%></font></td> 
    <td width="31%" align="center" height="18"><input type="button" name="<%=cstr(i)%>" value="删除" onClick="<%="DeleteIt(" + cstr(i)+ ")"%>"><input type="button" value="放入购物车" name="shop" onClick="<%="MoveToShop(" + cstr(i) + ")"%>"></td> 
  </tr> 
  <%
    next 
  %>
   
   
        
           
</table>       
      
<table  border="0" width="100%" bordercolorlight="#FFFFFF" cellspacing="0" cellpadding="0" bordercolordark="#808080">      
	<tr>      
		<td>      
          <p align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="button" name="clear" value="全部清除" onClick="Mybag.Clear.value='Y'; Mybag.action='favorite.asp'; Mybag.submit()">&nbsp;              
          <input type="button" value="继续购物" name="close" onClick="window.close()">                    
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
              
</table>                
</body>                  
                  
</html>                  
