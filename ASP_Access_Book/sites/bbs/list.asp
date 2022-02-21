<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="odbc_connection.asp"-->
<!--#include file="char.asp"-->
<!--#include file="code.asp"-->
<html>
<head>
<title>用户列表</title>
<!-- #include file="an.htm" -->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid; background-color: #e3f1d1;">
  <tr> 
    <td width="475" height="24"><a href="index.asp">论坛论坛</a>→<% if Trim(Request.QueryString("order"))="2" then %>
      <a href="list.asp?order=2">热门话题</a>
<% Else %>
      <a href="list.asp?order=1">论坛新贴</a>
<% End If %>
</td>
    <td width="283"> <a href="list.asp?order=1">论坛新贴</a> | <a href="list.asp?order=2">热门话题</a> 
      | <a href="uerlist.asp?order=2">发贴排名</a> | <a href="uerlist.asp">用户列表</a> |</td>
  </tr>
    <tr align="center"> 
    <td height="15" colspan="9" background="pic/h3.gif" class="p9"><strong><font color="#FFFFFF"><% if Trim(Request.QueryString("order"))="2" then %>热门话题  <% Else %>
论坛新贴<% End If %></font></strong></td>
  </tr>
</table>
<table width="760" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr align="center" class="p9"> 
    <td width="6%" background="pic/h3.gif">序号</td>
    <td width="55%" background="pic/h3.gif">标题</td>
    <td width="6%" background="pic/h3.gif">回复</td>
    <td width="6%" background="pic/h3.gif">点击</td>
    <td width="15%" background="pic/h3.gif">发言人</td>
    <td width="12%" background="pic/h3.gif">发言时间</td>
  </tr>
  <%
if Trim(Request.QueryString("order"))="2" then
sql = "SELECT * FROM news where layer=1 order by child desc,sorttime desc"
else
sql = "SELECT * FROM news where layer=1 order by sorttime desc"
end if
		dim sql,rs
		'因为要分页显示查询结果，所以用下面方法创建一个recordset对象
		set rs=Server.CreateObject("ADODB.Recordset")
		rs.Open sql,db,1,1          
		if not rs.bof and not rs.eof then
			'以下主要为了分页显示
			dim page_size                         '定义每页多少条记录变量
			dim page_no                           '定义当前是第几页变量
			dim page_total                        '定义总页数变量
			page_size=10                         '每页显示10条记录
			if Request("page_no")="" then         '如果第一次打开，则page_no为1，否则，由传 
				page_no=1                         '回的参数决定
			else
				page_no=cint(Request("page_no"))
			end if
			session("page_no")=page_no            '将page_no存入session,以备其它页返回时用
			rs.pagesize=page_size                 '设置每页多少条记录
			page_total=rs.pagecount               '返回总页数
			rs.absolutepage=page_no               '设置当前显示第几页
			'下面一段显示当前页的所有记录
			dim i,j
			i=0
			j=page_size                           '该变量用来控制显示多少条记录
			do while not rs.eof and j>0            '循环知道当前页结束或文件结尾
				i=i+1
				j=j-1
		%>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td><%=(page_no-1)*page_size+i%> 
    <td align="left" class="p9"><a href="count_hits.asp?bbs_id=<%=rs("bbs_ID")%>&boardid=<%=rs("boardid") %>"><%=ubbcode(RS("title"))%></a></td>
    <td class="p9"><%=RS("child")%></td>
    <td class="p9"><%=RS("hits")%></td>
    <td class="p9"><a href="uerparticle.asp?id=<%=server.HTMLEncode(rs("user_name"))%>"><%=rs("user_name")%></a></td>
    <td class="p9"><% sub_time=rs("submit_date")
	   Response.Write year(sub_time)&"-"&month(sub_time)&"-"&day(sub_time) %></td>
  </tr>
  <% 
		   rs.movenext
		   loop
		end if
		rs.close
		%>
</table>
<%
Set rs = Nothing
db.close
set db=nothing
%>
<!--#include file="foot.asp"-->

