<!--#include file="conn.asp"-->
<%dim key
key=replace(trim(request("key")),"'","")
if key="" then
response.write "<script>alert('请输入搜索关键字');window.location='index.asp';</script>"
end if%>
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
              <td width="590" colspan="2">　</td>
            </tr>
          </table>            <br>            <table width="560"  border="0" align="center" cellpadding="3" cellspacing="3" class="txt1">
            <%const MaxPerPage=20
                            dim totalPut   
   							dim CurrentPage
   							dim TotalPages
   							dim i,j
   							dim idlist
   							if not isempty(request("page")) then
   							currentPage=cint(request("page"))
   							else
   							currentPage=1
   							end if
                            Set rs= Server.CreateObject("ADODB.Recordset")
                            sql="select distinct articleid,title,addtime from article where (title like '%"&key&"%' or writer='"&key&"') and (classid<>19 and classid<>8 and classid<>2 and classid<>20 and classid<>7) order by addtime desc"
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
                              showpage totalput,MaxPerPage,"search.asp?key="&server.urlencode(key)
                            else
                              if (currentPage-1)*MaxPerPage<totalPut then
                                rs.move  (currentPage-1)*MaxPerPage
                                dim bookmark
                                bookmark=rs.bookmark
                                showContent
                                showpage totalput,MaxPerPage,"search.asp?key="&server.urlencode(key)
                              else
                                currentPage=1
                                showContent
                                showpage totalput,MaxPerPage,"search.asp?key="&server.urlencode(key)
                              end if
                            end if
			    else
				response.write "<tr><td width=100% colspan=3>没有找到 <font color=red>"&key&"</font> 相关资料，请重新搜索！</td></tr>"
                            end if
	        	    rs.close
                            set rs=nothing  
                            conn.close
                            set conn=nothing
                            sub showContent
                            i=0
                            %><tr><td width="100%" colspan="3">您搜索的关键字是：<font color=red><%=key%></font>，共找到结果 <font color=red><%=totalput%></font> 条</td>
            				</tr><%
                            do while not rs.eof and i<MaxPerPage
                            %><tr><td width="4%"><img src="images/iocn1.gif" width="7" height="7"></td>
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
                                   response.write "<a class=list3 href="&filename&"&page=1>首页</a>&nbsp;"
                                   response.write "<a class=list3 href="&filename&"&page="&CurrentPage-1&">上一页</a>&nbsp;"
                                end if
                                if n-currentpage<1 then
                                  response.write "<font color='#000080'>下一页 尾页</font>"
                                else
                                  response.write "<a class=list3 href="&filename&"&page="&(CurrentPage+1)&">"
                                  response.write "下一页</a> <a class=list3 href="&filename&"&page="&n&">尾页</a>"
                                end if 
                                response.write "<font color='#000080'>&nbsp;页次：</font><strong><font color=red>"&CurrentPage&"</font><font color='#000080'>/"&n&"</strong>页</font> "
                                response.write "<font color='#000080'>&nbsp;共<b>"&totalnumber&"</b>条 <b>"&maxperpage&"</b>条/页</font> "%></td>   <td class="txt3">　</td> </tr><%end function%>
                  				</table></td>
        </tr>
      </table>
      <!--#include file="end.inc"--></td>
  </tr>
</table>
</body>
</html>
