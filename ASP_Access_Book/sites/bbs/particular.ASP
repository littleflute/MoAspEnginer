<% Time1=Timer %>
<!--#include file="odbc_connection.asp"-->
<!--#include file="char.asp"-->
<!--#include file="code.asp"-->
<HTML>
	<head>
<% 	dim bbs_id
	bbs_id=request("bbs_id")     
Set rs = Server.CreateObject("ADODB.Recordset")        
rs.Open "Select * From  news where bbs_id =" & bbs_id, db,1,1 
'title=rs("title")
 %>
 		<title>详细内容--"<%=title%>"</title>
<!-- #include file="an.htm" -->
<!-- #include file="menu.asp" -->
<!--#include file="login_include.asp"-->
<!--#include file="menu_include.asp"-->
<!--#include file="menulist_include.asp"-->
<%       
	dim I                 
	I=0                                          
if (not rs.eof)  then                                        
		I=I+1          
	userid=rs("user_name")
	title=rs("title")    
	if rs("parent_id")<>0 then parent_id=rs("parent_id") else parent_id=0 end if
%>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="90"><a href="announce.asp?boardid=<%=boardid%>"><img src="pic/h4.gif" alt="发表文章"  title="发表文章" width="72" height="21" border="0"></a></td>
    <td width="542">
	<% If parent_id=0 Then %>
	<a href="re_announce.asp?bbs_id=<%=bbs_id%>&boardid=<%=boardid%>&title=<%=rs("title")%>"><img src="pic/h7.gif" width="72" height="21" border="0"></a>
	<% Else %>
	<a href="re_announce.asp?bbs_id=<%=parent_id %>&boardid=<%=boardid%>&title=<%=rs("title")%>"><img src="pic/h7.gif" width="72" height="21" border="0"></a>
	<% End If %>
</td>
    <td width="128"><img src="pic/team2.gif" width="20" height="20">版主：<%= banzhu %></td>
  </tr>
</table>
<table width="760"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr class="p9"> 
    <td width=20% background="pic/h3.gif">主题</td>
    <td background="pic/h3.gif"><b><%=ubbcode(title)%></b></td>
  </tr>
  <%
 Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open "Select * From uers where id ='" & userid & "'", db,1,1  

  %>
  <tr class="p9"> 
    <td rowspan="2" align="center" valign="top"> <% if rs1("face")<>"" then %> <img src="<%=rs1("face")%>"> <% else %> <img src="images/none.jpg"> <%end if %> <br>
      <table width="90%" border="0" cellspacing="0" cellpadding="0">
        <tr align="center"> 
          <td colspan="2"><a href="uerparticle.asp?id=<%=rs1("id")%>"><%=server.HTMLEncode(rs1("id"))%></a></td>
        </tr>
        <tr> 
          <td width="44%"><font color="#666666">用户头衔:</font></td>
          <td width="56%"><font color="#666666"><%=rs1("head")%></font></td>
        </tr>
        <tr> 
          <td><font color="#666666">用户等级:</font></td>
          <td><font color="#666666"><%=rs1("dengji")%></font></td>
        </tr>
        <tr> 
          <td><font color="#666666">发表文章:</font></td>
          <td><font color="#666666"><%=rs1("totle")%></font></td>
        </tr>
      </table>
      </td>
    <td valign=top>内容：<br> <%=ubbcode(rs("body"))%></td>
  </tr>
  <tr class="p9"> 
    <td align=right><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="4%">&nbsp;</td>
          <td width="12%"><% If session("AdminUID") = rs1("id") or session("Adminflag")= boardid  or session("Adminflag")="0" Then %> <a href="edit.asp?bbs_id=<%= bbs_id %>&boardid=<%= boardid %>">编辑</a> 
            <% Else %>
            编辑 <% End If %> </td>
          <td width="12%"><% If session("AdminUID") = rs1("id") or session("Adminflag")= boardid  or session("Adminflag")="0" Then %> <a href="delsave.asp?bbs_id=<%= bbs_id %>&boardid=<%= boardid %>">删除</a> 
            <% Else %>
            删除 <% End If %></td>
          <td width="12%"><% If (session("Adminflag"))= boardid  or session("Adminflag")="0" Then %> <a href="up.asp?bbs_id=<%= bbs_id %>&boardid=<%= boardid %>&title=<%=title %>">置顶</a> 
            <% Else %>
            置顶 <% End If %></td>
          <td>&nbsp;</td>
          <td width="16%">点击：<%=RS("hits")%></td>
          <td width="26%">时间：<%=rs("submit_date")%></td>
        </tr>
      </table></td>
  </tr>
</table>
  
<%
 rs1.close
 set rs1=nothing
 
 else
 
 Response.Write "<br><br><center>参数错误，不存在不文章！</center>"
 Response.End
 end if	

%>
<% 
Set rs = Server.CreateObject("ADODB.Recordset")        
rs.Open "Select * From  news where parent_id=" & bbs_id &"and layer = 2 order by bbs_id ", db,1,1 
do while (not rs.eof)       
 %>
<table width="760"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr class="p9"> 
    <td width=20% background="pic/bg_re.gif">回复</td>
    <td background="pic/bg_re.gif"><b><%=ubbcode(rs("title")) %></b></td>
  </tr>
  <% Set rs1 = Server.CreateObject("ADODB.Recordset")       
rs1.Open "Select * From uers where id ='" & rs("user_name") & "'", db,1,1       
  %>
  <tr class="p9"> 
    <td rowspan="2" align="center" valign="top"> <% if rs1("face")<>"" then %> <img src="<%=rs1("face")%>"> <% else %> <img src="images/none.jpg"> <%end if %> <br> <table width="90%" border="0" cellspacing="0" cellpadding="0">
        <tr align="center"> 
          <td colspan="2"><a href="uerparticle.asp?id=<%=rs1("id")%>"><%=server.HTMLEncode(rs1("id"))%></a></td>
        </tr>
        <tr> 
          <td width="44%"><font color="#666666">用户头衔:</font></td>
          <td width="56%"><font color="#666666"><%=rs1("head")%></font></td>
        </tr>
        <tr> 
          <td><font color="#666666">用户等级:</font></td>
          <td><font color="#666666"><%=rs1("dengji")%></font></td>
        </tr>
        <tr> 
          <td><font color="#666666">发表文章:</font></td>
          <td><font color="#666666"><%=rs1("totle")%></font></td>
        </tr>
      </table>
	  </td>
    <td valign=top>内容：<br> <%=ubbcode(rs("body"))%></td>
  </tr>
  <tr class="p9"> 
    <td align=right><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="4%">&nbsp;</td>
          <td width="12%"> <% If session("AdminUID") = rs1("id") or (session("Adminflag"))= boardid  or session("Adminflag")="0" Then %> <a href="edit.asp?bbs_id=<%= rs("bbs_id") %>&boardid=<%= boardid %>">编辑</a> 
            <% Else %>
            编辑 
            <% End If %> </td>
          <td width="12%"> <% If session("AdminUID") = rs1("id") or (session("Adminflag"))= boardid  or session("Adminflag")="0" Then %> <a href="delsave.asp?bbs_id=<%= rs("bbs_id")%>&boardid=<%= boardid %>">删除</a> 
            <% Else %>
            删除 
            <% End If %> </td>
          <td width="12%">&nbsp; </td>
          <td>&nbsp;</td>
          <td width="16%">点击：<%=RS("hits")%></td>
          <td width="26%">时间：<%=rs("submit_date")%></td>
        </tr>
      </table></td>
  </tr>
</table>
  
<%
		rs.movenext
	loop
 rs.close
 set rs=nothing	

%>
<a href="index.asp?page_no=<%=session("page_no")%>"></a>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="90"><a href="announce.asp?boardid=<%=boardid%>"><img src="pic/h4.gif" alt="发表文章"  title="发表文章" width="72" height="21" border="0"></a></td>
    <td width="125"><a href="re_announce.asp?bbs_id=<%=bbs_id%>&boardid=<%=boardid%>&title=<%=title %>">
	<% If parent_id=0 Then %>
	<a href="re_announce.asp?bbs_id=<%=bbs_id%>&boardid=<%=boardid%>&title=<%=title %>"><img src="pic/h7.gif" width="72" height="21" border="0"></a>
	<% Else %>
	<a href="re_announce.asp?bbs_id=<%=parent_id %>&boardid=<%=boardid%>&title=<%=title %>"><img src="pic/h7.gif" width="72" height="21" border="0"></a>
	<% End If %>
    </td>
    <td width="106"><a href="board.asp?page_no=<%=session("page_no")%>&boardid=<%=boardid%>">返回文章列表</a> 
    </td>
    <td width="268">
<% If parent_id<>0 Then %>
 <a href="particular.asp?bbs_id=<%= parent_id %>&boardid=<%= boardid %>">同主题阅读</a>	
<% Else %>同主题阅读<% End If %>	
	</td>
    <td width="171">&nbsp;</td>
  </tr>
</table>
<a href="index.asp?page_no=<%=session("page_no")%>"> </a>
<% Time2=Timer %>
<!--#include file="foot.asp"-->