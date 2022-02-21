<%
	dim Nstr(20)
	
	Total=Request.Cookies("times") 
	if cstr(Total) = 0 then 
		response.write "您的购物车中没有任何商品!"
		response.end
	end if
	for i = 1 to Clng(Total) 
		Nstr(i)=Request.Cookies("MYShopBag"+cstr(i)) 
	next	 

	if Request.Cookies("H19681019S19681231W19980412B69793682") <> "OK" then
		response.redirect "member.asp"
	end if
%>


<html>
<head>
<title>收银台</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="ditu01.gif"><form name="Transfer" method="POST" action="MailToMe.asp">

<table width="520" border="0" cellspacing="0" cellpadding="0" align="center" height="451">
  <tr valign="bottom"> 
      <td width="21" height="38"><img src="image/money01.gif" width="21" height="59"></td>
      <td width="200" background="image/buyditu01.gif" height="38"><img src="image/money02.gif" width="57" height="59"></td>
      <td align="right" background="image/buyditu01.gif" width="418" height="38"><img src="image/money03.gif" width="185" height="35"></td>
      <td align="right" width="35" height="38"><img src="image/buy04.gif" width="35" height="35"></td>
  </tr>
  <tr> 
      <td width="21" background="image/buyditu04.gif" height="342">&nbsp;</td>
    <td colspan="2" valign="top" height="342"> 
      <table width="98%" border="0">
        <tr> 
          <td><font size="2"><span  class="hongzi">1.下面是您的购物订单，请仔细核对商品名称、数量及价格</span></font></td>
        </tr>
      </table>
      <table width=98% border=0 cellspacing=1 cellpadding=0 >
        <tr align=center valign=bottom bgcolor="#666666"> 
          <td height=18 width=120><font color="#FFFFFF" size="2">商品名称</font></td>
          <td height=18 width=70><font color="#FFFFFF" size="2">会员价</font></td>
          <td height=18 width=60><font color="#FFFFFF" size="2">数量</font></td>
            <td height=18 width=80><font color="#FFFFFF" size="2">会员价小计</font></td>
        </tr>
        
  <%
  	dim S
  	dim CC
  	
  	for i = 1 to Total 
  		S=split(Nstr(i),"/")
  		CC=CC+Clng(S(1))*Clng(S(2))
  %> 
  
        <tr align=center valign=bottom> 
          <td height=18 width=120><font color="#0000cc" size="2"><%=S(0)%></font></td>
          <td height=18 width=70><font color="#CC0000" size="2"><%=S(1)%></font></td>
          <td height=18 width=60><font size="2"><%=S(2)%></font></td>
          <td height=18 width=80><font color="#CC0000" size="2"><%=cstr(clng(S(1))*clng(S(2)))%></font></td>
        </tr>
       
  <%
    next 
  %>
       
       
      </table>
      <hr width=100%>
      <table width=98% border=0 cellspacing=1 cellpadding=0>
        <tr align=center valign=bottom> 
          <td height=18 width=120>&nbsp;</td>
          <td height=18 width=70>&nbsp;</td>
          <td height=18 width=70>&nbsp;</td>
          <td height=18 width=60><font color="#0000cc" size="2"> 总计： </font></td>
          <td height=18 width=80><font color="#CC0000" size="2"> <%=CC%>元</font></td>
        </tr>
      </table>
<p><font  size="2">&nbsp; &nbsp; ★如果你要打印此订单，请单击鼠标右键并从弹出的菜单中选择“打印”命令。<br>&nbsp; &nbsp; ☆如果要保存此订单，请单击鼠标右键后选择“查看源文件”命令，再执行“文件”/“另存为”命令，将文件存成 .htm文件放到你的硬盘中即可。</font></p>
    <table width="98%" border="0">
        <tr>
          <td><font size="2">2.送货的方式</font></td>
        </tr>
      </table>
      <table cellspacing=0 cellpadding=0 width=100% border=1 style='font-size: 9pt' bordercolorlight='#D2D2D2' bordercolordark='#FFFFFF' align=center height="42">
        <tr> 
          <td height="18">
            <table border=0 width="442" >
              <tr> 
                <td height="29" width="73" valign="top"> <font size="2"> 
                  <input type='radio' name='Sendit' value='平邮信函' checked >
                  <b>平邮</b></font></td>    
                <td height="29" width="355"><font size="2">我们将以挂号印刷品或包裹的方式发送。在零售价的基础上,我们免费邮寄。或在批发价的基础上,收取附加费<font color="#0000FF">10元</font>（挂号1元＋保险1元＋包装材料4元＋实际邮资）。</font></td>    
              </tr>    
            </table>    
          </td>    
        </tr>    
      </table>    
      <table cellspacing=0 cellpadding=0 width="98%">    
        <tr>     
          <td align=left width="80%" class="hongzi"><font size="2">3.请选择付款方式</font></td>    
        </tr>    
      </table>    
      <table cellspacing=0 cellpadding=0 width=100% border=1 style='font-size: 9pt' bordercolorlight='#D2D2D2' bordercolordark='#FFFFFF' align=center height="131">    
        <tr>     
          <td height="130">     
            <table border=0  width="445">    
              <tr>     
                <td valign=top width="87" align="left">  
                  <p align="left"> <font size="2">    
                  <b><input type="checkbox" name="C1" value="ON" checked> 款到付货</b></font></p>
                </td>  
                <td width="344"><font size="2"><null> 1.<font color="#0000FF">通过邮局汇款</font>，我们收到汇款单一般要3-5个工作日(偏远的地区还会更慢些)。<br>
 2.<font color="#0000FF">通过银行电汇</font>，到帐要3个工作日左右(偏远的地区还会更慢些)。收到汇款后,我们将会及时给您发货。请您电汇时尽量保证汇款人名称和订单收货人名称一致，因为有时候即使您留下了订单号也无法在电汇单上显示出来，这样可能无法查到您的订单号，会耽搁您的订单发货。<br>
&nbsp;&nbsp;<font color="#FF0000">注意：</font><font color="#48A289">北京地区四环以内的用户可送货上门，但要加5元的配送费。</font>
</font></td>  
              </tr>  
              <tr>   
                <td valign=top width="87"> <font size="2">   
                  <input type='radio' name='Recieve' value='1' checked >  
                  <b>邮局汇款</b></font></td>    
                <td width="344"><font size="2">请在汇款人简短留言中注明您的<font color="#0000FF">订单号</font>/<font color="#0000FF">用户名</font>（很重要） 
                  <br>
                    汇款地址：北京市xxxxx<br>
                    邮编：100000<br>
                    电话：010-xxxxxx</font></td>    
              </tr>    
              <tr>     
                <td valign=top width="87"> <font size="2">     
                  <input type='radio' name='Recieve' value='2' >    
                  <b>银行电汇</b></font></td>                             
                <td width="344"><font size="2">公司：郑州亮中<br>
                    开户行：xxxx支行<br>
                    帐号：xxxxxxx</font></td>                            
              </tr>                            
            </table>                            
          </td>                            
        </tr>                            
      </table>                            
          </td>                                   
      <td width="35" background="image/buyditu03.gif" height="342">&nbsp;</td>                                   
  </tr>                                   
  <tr valign="top">                                    
      <td width="21" height="27"><img src="image/buy05.gif" width="21" height="21"></td>                  
      <td colspan="2" background="image/buyditu02.gif" height="27">                   
        <p align="right">                  
<input type="button" value="让我再想想?" name="back" onClick="history.go(-1)">&nbsp;<input type="submit" value="发送请求" name="Send"></p>                     
    </td>                                 
      <td width="35" height="27"><img src="image/buy06.gif" width="35" height="21"></td>                                 
  </tr>                                 
  <tr>                                  
    <td width="21" height="21">&nbsp;</td>                                 
    <td width="200" height="21">&nbsp;</td>                                 
    <td width="418" height="21">&nbsp;</td>                                 
    <td width="35" height="21">&nbsp;</td>                                 
  </tr>                                 
</table>                                 
<table width="520" border="0" cellspacing="0" cellpadding="0" align="center">                          
  <tr>                                 
    <td width="768" height="4" align="center">                                  
      <hr width="100%" size="1" noshade>                                 
    </td>                                 
  </tr>                                 
  <tr>                                  
    <td width="768" height="4" align="center"><font size="2"><b>Copyright</b>:郑州亮中</font></td>                                 
  </tr>                                 
  <tr>                                  
    <td width="768" height="4" align="center"><font size="2"><br>          
      E-mail:<a href="mailto:xxxx@xxxx.com">xxxx@xxxx.com</a></font></td>                              
  </tr>                              
</table>                   
</form>                             
</body>                              
</html>                              
