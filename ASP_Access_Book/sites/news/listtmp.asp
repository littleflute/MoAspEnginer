<!--#include file="conn.asp"-->
<%dim classid
classid=replace(trim(request("classid"))," ","")
if classid="" then
response.redirect "index.asp"
end if%>
<html>
<head>
<title>������</title>
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
              <td width="590" colspan="2"><img src="images/deng<%=classid%>.gif" width="580" height="60"></td>
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
                              'showpage totalput,MaxPerPage,"list.asp?classid="&classid
                              showContent
                              showpage totalput,MaxPerPage,"list.asp?classid="&classid
                            else
                              if (currentPage-1)*MaxPerPage<totalPut then
                                rs.move  (currentPage-1)*MaxPerPage
                                dim bookmark
                                bookmark=rs.bookmark
                                'showpage totalput,MaxPerPage,"list.asp?classid="&classid
                                showContent
                                showpage totalput,MaxPerPage,"list.asp?classid="&classid
                              else
                                currentPage=1
                                'showpage totalput,MaxPerPage,"list.asp?classid="&classid
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
                               end if%><tr><td>��</td>
                               <td align="center"><%
                                if CurrentPage<2 then
                                   response.write "<font color='#000080'>��ҳ ��һҳ</font>&nbsp;"
                                else
                                   response.write "<a href="&filename&"&page=1 class=list3>��ҳ</a>&nbsp;"
                                   response.write "<a href="&filename&"&page="&CurrentPage-1&" class=list3>��һҳ</a>&nbsp;"
                                end if
                                if n-currentpage<1 then
                                  response.write "<font color='#000080'>��һҳ βҳ</font>"
                                else
                                  response.write "<a href="&filename&"&page="&(CurrentPage+1)&" class=list3>"
                                  response.write "��һҳ</a> <a href="&filename&"&page="&n&" class=list3>βҳ</a>"
                                end if 
                                response.write "<font color='#000080'>&nbsp;ҳ�Σ�</font><strong><font color=red>"&CurrentPage&"</font><font color='#000080'>/"&n&"</strong>ҳ</font> "
                                response.write "<font color='#000080'>&nbsp;��<b>"&totalnumber&"</b>�� <b>"&maxperpage&"</b>��/ҳ</font> "%></td>   <td class="txt3">��</td> </tr><%end function%>
                  				</table></td>
        </tr>
      </table>
      <!--#include file="end.inc"--></td>
  </tr>
</table>
</body>
</html>
