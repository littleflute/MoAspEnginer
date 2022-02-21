<!--#include file="odbc_connection.asp"-->
<!--#include file="function2.asp"-->
<!--#include file="check.asp"-->
<html>
	<head>
		<title>小论坛</title>
		
<!-- #include file="an.htm" -->
<!-- #include file="menu.asp" -->
<% If (not session("AdminUID")="") and session("Adminflag")="0" Then %>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid; background-color: #e3f1d1;">
    <td width="469" height="16"><a href="index.asp">论坛</a> →<a href="oneedit.asp">超级管理</a> 
      →→ <a href="sql.asp">执行sql语句</a> | </td>
    <td width="289">管理菜单：| <a href="board.asp?boardid=0">查看公告</a> <a href="oneedit.asp?pub=yes">| 
      管理公告 </a> | <a href="announce.asp?boardid=0">发布公告 </a> |</td>
  </tr>
</table>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="27" class="p9"> <p> 
        <% =session("AdminUID") %>
        ，<font color="#FF0000">请务必注意！！！如果要删除，一定要看清楚。现在还不能修改，只能删除。删除后将不能取消！！！<br>
        另: 请特别注意：如果删除的是主题，那么这个主题的所有回复都将被删除！！如果是回复的文章，将只删除此回复，主题不受影响</font></p></td>
  </tr>
</table>
<table width="760" border="0" align="center" cellspacing="2" bordercolor="#8800FF" >
  <tr align="center" class="p9"> 
    <td width="5%"  background="pic/h3.gif">序号</td>
    <td  background="pic/h3.gif">主题</td>
    <td width="10%" background="pic/h3.gif">版面</td>
    <td width="5%" background="pic/h3.gif">回复</td>
    <td width="9%" background="pic/h3.gif">发言人</td>
    <td width="15%" background="pic/h3.gif">发言时间</td>
    <td width="11%" background="pic/h3.gif">IP</td>
    <td width="5%" background="pic/h3.gif">删除</td>
  </tr>
  <%
		dim sql,rs
		'因为要分页显示查询结果，所以用下面方法创建一个recordset对象
	    if request("pub")="yes" then
		sql="select * from news where boardid= 0 order by submit_date desc"
		else	
		sql="select * from news  order by submit_date desc"
		end if
		set rs=Server.CreateObject("ADODB.Recordset")
		rs.Open sql,db,1,1            
		if not rs.bof and not rs.eof then
			'以下主要为了分页显示
			dim page_size                         '定义每页多少条记录变量
			dim page_no                           '定义当前是第几页变量
			dim page_total                        '定义总页数变量
			page_size=20                          '每页显示10条记录
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
			do while not rs.eof and j>0           '循环知道当前页结束或文件结尾
				i=i+1
				j=j-1
				board_id=RS("boardid")
		%>
  <tr align="center"> 
    <td height="20" bgcolor="#FFFFCC"><%=RS("bbs_id")%> 
    <td bgcolor="#FFFFCC" class="p9"><a href="count_hits.asp?bbs_id=<%=rs("bbs_ID")%>&boardid=<%= board_id %>"><%=RS("title")%></a></td>
    <td bgcolor="#FFFFCC" class="p9"> <%
sqlboardid="select * from board where boardid="&board_id
set rsboardid= db.execute(sqlboardid)
%> <a href="board.asp?boardid=<%=board_id%>"><%= rsboardid("name") %></a> <%
rsboardid.close
set rsboardid=nothing
%> </td>
    <td bgcolor="#FFFFCC" class="p9"><%=RS("child")%></td>
    <td bgcolor="#FFFFCC" class="p9"><a href="uerparticle.asp?id=<%=rs("user_name")%>"><%=rs("user_name")%></a></td>
    <td bgcolor="#FFFFCC" class="p9"><%=rs("submit_date")%></td>
    <td bgcolor="#FFFFCC" class="p9"><%=rs("ip")%></td>
    <td bgcolor="#FFFFCC" class="p9"><a href="del.asp?bbs_id=<%=RS("bbs_id")%>"><font color="#FF0000">删除</font></a></td>
  </tr>
  <%
				rs.movenext
			loop
		end if
		%>
</table>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center"> 
      <%
	'调用子程序，写出有关各页的链接信息
	call select_page(page_no,page_total)
	%> </td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
</table>

<% Else %>

<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="center">&nbsp;你没有管理权限或请重新登陆！<a href="javascript:history.go(-1)">请点击这里返回！</a> 
	</td>
  </tr>
</table>
<% End If %>
<!--#include file="foot.asp"-->
