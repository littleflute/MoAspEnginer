
<html>

<head>
<title>首页更新</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet>
</head>
<body>
<table width="83%" height="170" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr><td><table width="103%" border="0" align="center" cellpadding="1" cellspacing="0" bordercolor="#000000" bordercolorlight="#999999" bordercolordark="#FFFFFF">
        <tr align=center>
          <td bgcolor="#CCCCCC"><B><font color="#000000">添加首页公告</font></B></td>
        </tr>
        <tr>
          <td> <form method="POST" action="save.asp" >
              <div align="center">标题 
                <input type="text" name="name" size="60" value="">
                <br>
                内容 <br>
                <textarea name="neirong" rows="3" cols="60"></textarea>
                <br>
                <input type="submit" name="Submit" value="提交">
                <input type="reset" name="Submit2" value="清除">
              </div>
            </form></td>
        </tr>
      </table></td></tr></table>
<%
If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
CurPage = 1
Else
CurPage = CINT(Request.QueryString("CurPage"))
End If
%>
<div align="center">
  <!-- #include file="connect.asp" -->
  <%rs.Open "Select * From data Order By time DESC", conn, 1,1%>
  <%if rs.eof and rs.bof then%>
  暂时没有任何记录！！ 
  <%else
RS.PageSize=10'设置每页记录数           
Dim TotalPages              
TotalPages = RS.PageCount              
              
If CurPage>RS.Pagecount Then               
CurPage=RS.Pagecount              
end if              
              
RS.AbsolutePage=CurPage
rs.CacheSize = RS.PageSize'设置最大记录数  
Dim Totalcount              
Totalcount =INT(RS.recordcount)              
              
StartPageNum=1              
do while StartPageNum+10<=CurPage              
StartPageNum=StartPageNum+10              
Loop              
                  
EndPageNum=StartPageNum+9              
              
If EndPageNum>RS.Pagecount then EndPageNum=RS.Pagecount       %>
</div>
<table border="0" width="82%" cellpadding="1" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#999999" align="center">
  <tr bgcolor="#CCCCCC"> 
    <td width="52%" height="20"><b>&nbsp;<font color="#000000">标题</font></b></td>
    <td width="35%" height="20"  align="center"><b><font color="#000000">时 间</font></b></td>
    <td width="13%" height="20" align="center"><b><font color="#000000">操作</font></b></td>
  </tr>
  <%I=0
p=RS.PageSize*(Curpage-1)              
do while (Not RS.Eof) and (I<RS.PageSize)              
p=p+1%>
  <tr> 
    <td bordercolorlight="#C0C0C0" align=left><b><%=rs("name")%></b></td>
    <td  align="center">[<%=rs("time")%>]</td>
    <td bordercolorlight="#c0c0c0" align=middle><a href="del.asp?id=<%=rs("id")%>">删除</a>|<a href="modifile.asp?id=<%=rs("id")%>">修改</a></td>
  </tr>
  <%I=I+1              
RS.MoveNext              
Loop%>
  <tr>
    <td colspan=4 align=middle bordercolorlight="#C0C0C0">页次： <font color="#CC0000"><%=CurPage%></font>/<%=TotalPages%> 
      页数： <a href="add.asp?CurPage=<%=StartPageNum-1%>">&lt;&lt;</a> 
      <% For I=StartPageNum to EndPageNum                
      if I<>CurPage then %>
      <a href="add.asp?CurPage=<%=I%>"><%=I%></a> 
      <% else %>
      <%=I%> 
      <% end if %>
      <% Next %>
      <% if EndPageNum<RS.Pagecount then %>
      <a href="add.asp?CurPage=<%=EndPageNum+1%>">下页</a> 
      <%end if%>
      |<a href="add.asp">刷 新</a></td>
  </tr>
</table>
<p align="center">　</p>
</body>
</html>
<%
end if
rs.close
set rs=nothing
%> 