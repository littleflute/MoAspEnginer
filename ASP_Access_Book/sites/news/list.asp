<!--#include file="conn.asp"-->
<%dim classid
classid=replace(trim(request("classid"))," ","")
if classid="" then
response.redirect "index.asp"
end if
tclassid=classid
set rs=conn.execute("select parentid,layer from class where classid="&classid)
if rs(1)>1 then
tclassid=rs(0)
end if
rs.close
%>
<html>
<head>
<title>新闻网</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="images/css.css" rel=stylesheet type=text/css>
<style type="text/css">
<!--
body {
	background-image: url(images/pagebg1.gif);
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
</head>

<body>
<table width="780" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CACACA">
  <tr>
    <td bgcolor="#FFFFFF"><!--#include file="top.inc"-->
    <table width="780" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="171" valign="top" bgcolor="F7EEE4"><!--#include file="left.inc"--></td>
          <td width="609" valign="top"><table width="580" border="0" align="center" cellpadding="5" cellspacing="5">
            <tr>
              <td width="590" colspan="2"><img src="images/deng<%=tclassid%>.gif" width="580" height="60"></td>
            </tr>
          </table>
          <%dim MaxPerPage
		dim bookmark
		dim totalPut   
   			    dim CurrentPage
   			    dim TotalPages
   			    dim i,j
   			    dim idlist
			 dim m
   			    if not isempty(request("page")) then
   			    	currentPage=cint(request("page"))
   			    else
   			    	currentPage=1
   			    end if
                            Set rs= Server.CreateObject("ADODB.Recordset")
	  set rs=conn.execute("select child from class where classid="&classid)
          dim classchild
          classchild=rs(0)
	  if classchild>6 then
		m=4
	  else
		m=int(20/classchild)
	  end if
          rs.close
          if classchild>1 then
	  MaxPerPage=6
          sql="select classid,class from class where parentid="&classid&" order by rootid,ordersid"
	  rs.open sql,conn,1,1
	  if not rs.eof then
                            totalPut=rs.recordcount
                            if currentpage<1 then
                              currentpage=1
                            end if
                            if (currentpage-1)*MaxPerPage>totalput then
                              if (totalPut mod MaxPerPage)=0 then
                                currentpage= totalPut \ MaxPerPage
                              else
                                 currentpage= totalPut \ MaxPerPage + 1
                              end if
                            end if
                            
                            if currentPage=1 then
                              showContent2
                              showpage2 totalput,MaxPerPage,"list.asp?classid="&classid
                            else
                              if (currentPage-1)*MaxPerPage<totalPut then
                                rs.move  (currentPage-1)*MaxPerPage
                                
                                bookmark=rs.bookmark
                                showContent2
                                showpage2 totalput,MaxPerPage,"list.asp?classid="&classid
                              else
                                showContent2
                                showpage2 totalput,MaxPerPage,"list.asp?classid="&classid
                              end if
                            end if
                            rs.close
                            end if
	        
                            set rs=nothing  
                            conn.close
                            set conn=nothing
     		sub showContent2
                            i=0
          	do while not rs.eof and i<MaxPerPage%>
			<div align="center">
				<table border="0" style="border-collapse: collapse" width="569" cellpadding="0" id="table1">
					<tr>
						<td height="35" background="images/lm.gif"><table width=100%><tr><td>&nbsp;&nbsp;&nbsp; <a href="list.asp?classid=<%=rs(0)%>" class="list3"><%=rs(1)%></a></td><td align=right><a href="list.asp?classid=<%=rs(0)%>" class="list3">更多 ...</a></td></tr></table>
						</td>
					</tr>
					<tr>
						<td><table width="560"  border="0" align="center" cellpadding="3" cellspacing="3" class="txt1" id="table1">
            				<%set rs1=conn.execute("select top "&m&" articleid,title,addtime from article where classid="&rs(0)&" order by addtime desc")
            				i=0
            				do while not rs1.eof and i<m%>
            				<tr><td width="4%"><img src="images/iocn1.gif" width="8" height="7"></td>
                            <td width="77%" class="txt1"><a href="show.asp?id=<%=rs1(0)%>" target="_blank" class="list3"><%=rs1(1)%></a> <%if month(rs1(2))=month(date()) and day(rs1(2))-day(date())>3 then%><img src="images/news.gif" width="30" height="13" align="absmiddle"><%end if%></td>
              				<td width="19%" class="txt3">[<%=year(rs1(2))&"-"&month(rs1(2))&"-"&day(rs1(2))%>]</td>
            				</tr>
            				<%rs1.movenext
            				loop
            				rs1.close%>
            				</table></td>
					</tr>
				</table>
			</div>
			<%rs.movenext
			loop
		end sub
			function showpage2(totalnumber,maxperpage,filename)
                               dim n
                               if totalnumber mod maxperpage=0 then
                                  n= totalnumber \ maxperpage
                               else
                                  n= totalnumber \ maxperpage+1
                               end if%><table width=100%><tr><td>　</td>
                               <td align="center"><%
                                if CurrentPage<2 then
                                   response.write "<font color='#000080'>首页 上一页</font>&nbsp;"
                                else
                                   response.write "<a href="&filename&"&page=1 class=list3>首页</a>&nbsp;"
                                   response.write "<a href="&filename&"&page="&CurrentPage-1&" class=list3>上一页</a>&nbsp;"
                                end if
                                if n-currentpage<1 then
                                  response.write "<font color='#000080'>下一页 尾页</font>"
                                else
                                  response.write "<a href="&filename&"&page="&(CurrentPage+1)&" class=list3>"
                                  response.write "下一页</a> <a href="&filename&"&page="&n&" class=list3>尾页</a>"
                                end if 
                                response.write "<font color='#000080'>&nbsp;页次：</font><strong><font color=red>"&CurrentPage&"</font><font color='#000080'>/"&n&"</strong>页</font> "
                                response.write "<font color='#000080'>&nbsp;共<b>"&totalnumber&"</b>条专题 <b>"&maxperpage&"</b>条/页</font> "%></td>   <td class="txt3">　</td> </tr></table><%end function		
			else%>
			<br>            <table width="560"  border="0" align="center" cellpadding="3" cellspacing="3" class="txt1">
            <%MaxPerPage=20
                            
                            sql="select articleid,title,addtime from article where classid="&classid&" order by addtime desc"
                            rs.open sql,conn,1,1
                            if not rs.eof then
                            totalPut=rs.recordcount
                            if currentpage<1 then
                              currentpage=1
                            end if
                            if (currentpage-1)*MaxPerPage>totalput then
                              if (totalPut mod MaxPerPage)=0 then
                                currentpage= totalPut \ MaxPerPage
                              else
                                 currentpage= totalPut \ MaxPerPage + 1
                              end if
                            end if
                            
                            if currentPage=1 then
                              showContent
                              showpage totalput,MaxPerPage,"list.asp?classid="&classid
                            else
                              if (currentPage-1)*MaxPerPage<totalPut then
                                rs.move  (currentPage-1)*MaxPerPage
 
                                bookmark=rs.bookmark
                                showContent
                                showpage totalput,MaxPerPage,"list.asp?classid="&classid
                              else
                                showContent
                                showpage totalput,MaxPerPage,"list.asp?classid="&classid
                              end if
                            end if
                            rs.close
                            end if
	        
                            set rs=nothing  
                            conn.close
                            set conn=nothing
                            sub showContent
                            i=0
                            do while not rs.eof and i<MaxPerPage
                            %><tr><td width="4%">
				<img src="images/iocn1.gif" width="8" height="7"></td>
                            <td width="77%" class="txt1"><a href="show.asp?id=<%=rs(0)%>" target="_blank" class="list3"><%=rs(1)%></a> <%if month(rs(2))=month(date()) and day(rs(2))-day(date())>3 then%><img src="images/news.gif" width="30" height="13" align="absmiddle"><%end if%></td>
              				<td width="19%" class="txt3">[<%=year(rs(2))&"-"&month(rs(2))&"-"&day(rs(2))%>]</td>
            				</tr><%i=i+1
                              if i>=MaxPerPage then exit do
                              rs.movenext
                              loop
                              end sub
                              function showpage(totalnumber,maxperpage,filename)
                               dim n
                               if totalnumber mod maxperpage=0 then
                                  n= totalnumber \ maxperpage
                               else
                                  n= totalnumber \ maxperpage+1
                               end if%><tr><td>　</td>
                               <td align="center"><%
                                if CurrentPage<2 then
                                   response.write "<font color='#000080'>首页 上一页</font>&nbsp;"
                                else
                                   response.write "<a href="&filename&"&page=1 class=list3>首页</a>&nbsp;"
                                   response.write "<a href="&filename&"&page="&CurrentPage-1&" class=list3>上一页</a>&nbsp;"
                                end if
                                if n-currentpage<1 then
                                  response.write "<font color='#000080'>下一页 尾页</font>"
                                else
                                  response.write "<a href="&filename&"&page="&(CurrentPage+1)&" class=list3>"
                                  response.write "下一页</a> <a href="&filename&"&page="&n&" class=list3>尾页</a>"
                                end if 
                                response.write "<font color='#000080'>&nbsp;页次：</font><strong><font color=red>"&CurrentPage&"</font><font color='#000080'>/"&n&"</strong>页</font> "
                                response.write "<font color='#000080'>&nbsp;共<b>"&totalnumber&"</b>条 <b>"&maxperpage&"</b>条/页</font> "%></td>   <td class="txt3">　</td> </tr><%end function%>
                  				</table>
               <%end if%></td>
        </tr>
      </table>
      <!--#include file="end.inc"--></td>
  </tr>
</table>
</body>
</html>
