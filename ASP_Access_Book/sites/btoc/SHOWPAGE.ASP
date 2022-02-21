<%
Sub ShowOnePage( rs, Page )
 
	
  	if Page = 0  then 
		response.write "<script language=""javascript"">"
		response.write "{"
		response.write "alert('嗨,真奇怪,居然没找到？！');"
		response.write "history.go(-1);"
		response.write "}"
		response.write "</script>"
		response.end
	end if

  	rs.AbsolutePage = Page

  	For iPage = 1 To rs.PageSize
		Ostr="window.open('introduce.asp?P=" + rs.fields("名称") + "&T=" + rs.fields("图标") + "',null,'width=520,height=400,scrollbars=1')"   
		Mstr="window.open('shopbag.asp?P=" + rs.fields("名称") + "&PP=" + rs.fields("定价") + "&Price=" + rs.fields("售价") + "',null,'width=560,height=360,scrollbars=1,status=1')"   
		Lstr="window.open('favorite.asp?P=" + rs.fields("名称") + "&Price=" + rs.fields("售价") + "&T=" + rs.fields("图标") + "',null,'width=560,height=360,scrollbars=1')"                                
		
    	RecNo = (Page - 1) * rs.PageSize + iPage

  		response.write "<table width=""100%"" border=""0"" cellspacing=""0"">"
  		response.write "<tr>"
  		response.write "    <td rowspan=""4"" valign=""top"" width=""16%""><a href onClick=""" + Ostr + """><img src=""../pro-image/" + rs.fields("图标") + """ width=""70"" height=""90""></a></td>"                               
  		response.write "    <td width=""84%""><font size=""2"" color=""#660000""><b><font color=""#000066"">商品名:</font><font size=""2"" color=""#990000""><b><font color=""#000066"">"
  		response.write rs.fields("名称") + "</font></b></font><font color=""#000066"">(</font><font size=""2"" color=""#990000""><b><font color=""#000066"">" + rs.fields("包装") + "</font></b></font><font color=""#000066"">)</font></b></font></td>"                               
  		response.write "</tr>"
  		response.write "<tr>"
  		response.write "<td width=""84%""><font size=""2"" color=""#660000""><u>原始价格:</u></font><font size=""2""><u><font color=""#660000""><strike>" + rs.fields("定价") + "</strike></font></u></font><font size=""2"" color=""#660000""><u>元 "
  		response.write "会员价格:</u></font><font size=""2""><u><font color=""#660000"">" + rs.fields("售价") + "</font></u></font><font size=""2"" color=""#660000""><u>元</u></font></td>" 
  		response.write "</tr>"
  		response.write "<tr>"
  		response.write "<td width=""84%""><br><font size=""2"">" + rs.fields("简介") + "</font></td>"
  		response.write "</tr>"
  		response.write "<tr>"
  		response.write "<td width=""84%"">"
  		response.write "<p align=""right""><font size=""2""> "
  		response.write "<a href style=""CURSOR: hand"" onClick=""" + Ostr + """><font color=""#0000FF"">详细资料 </font></a>"
  		response.write "<img border=""0"" src=""image/cart.gif""><a href style=""CURSOR: hand"" onClick=""" + Mstr + """><font color=""#0000FF"">购买 </font></a>"
  		response.write "<img border=""0"" src=""image/favor.gif""><a href style=""CURSOR: hand"" onClick=""" + Lstr + """><font color=""#0000FF"">收藏</font></a> "
  		response.write "</font></p>"
  		response.write "</td>"
  		response.write "</tr>"
  		response.write "<td width=""16%""><center>"
  		response.write "<tr>"
  		response.write "<td colspan=""2""><img src=""indeximage/line.gif"" width=""573"" height=""10""></td> "
  		response.write "</tr>"
  		response.write "</table>"
  		  		
  		rs.MoveNext
   	If rs.EOF Then Exit For
  Next
End Sub
%>

                                                                   
                                                                           
                                                                           
                                                                           
                                                                          
                                                                       
<p></p>
