<%@ language=vbscript%>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="guest_sub.asp"-->
<%Header"---回复留言"%>

<head>
<title></title>
</head>

<body>
 <center>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
 <IMG SRC="images/list.gif" BORDER="0" WIDTH="16" HEIGHT="16" ><a href="guest_list.asp?PageNo=1">查看以往留言</a></center> 
 <table border=0 cellspacing=0 cellpadding=0  width=455 align=center bordercolor=#9E9E9E> 
   <td align=center> 
    <br> 
	<form method="POST" action="guest_mreply.asp"> 
    <input type="hidden" name="from" value="SaveReply"> 
    <input type="hidden" name="PageNo" value="<%=Request.QueryString("PageNo")%>"> 
    <input type="hidden" name="ID" value="<%=Request.QueryString("ID")%>"> 
    <table border=0 cellspacing=0 cellapdding=0 align=center width=453> 
     <tr> 
	  <td width=105 align=center><IMG SRC="images/passwd.gif" BORDER="0" WIDTH="16" HEIGHT="16" > 
      &nbsp;管理密码</td>
	  <td width="340"><input type="text" name="PassWord" size="40" style="background-color: #FFFFFF; border-style: solid; border-width: 1"></td>
	 </tr>
	 <tr><td height=10 width="105"></td></tr>
	 <tr>
	  <td width=105 align=center><IMG SRC="images/msglist.gif" BORDER="0" WIDTH="16" HEIGHT="21" >
         &nbsp;回复主题
	  </td>
	  <td height=10 width="340"><input type="text" name="reply" size="40" style="background-color: #FFFFFF; border-style: solid; border-width: 1"></td>
	  </tr>
	<tr>
	  <td align=center width="105"><IMG SRC="images/reply.gif" BORDER="0" WIDTH="16" HEIGHT="16" >
         &nbsp;回复文章</td>
	  <td width="340"><textarea rows="7" name="Content" cols="40" style="background-color: #FFFFFF; border-style: solid; border-width: 1"></textarea></td>
	</tr>
    </table>
	<br>
	<center><input type="submit" value=" 提  交 " name="B1" style="background-color: #C0C0C0; font-family: 宋体; font-size: 10pt; border: 1 solid #FFFFFF">
    <input type="reset" value=" 清  除 " name="B2" style="background-color: #C0C0C0; font-family: 宋体; font-size: 10pt; border: 1 solid #FFFFFF"> 
    <br> 
 </form> 
</table> 
<br><br> 
<center><%Bottom%></center> 
