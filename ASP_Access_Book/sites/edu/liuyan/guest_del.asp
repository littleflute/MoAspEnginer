<!--#INCLUDE FILE="guest_sub.asp" -->
<%
Header "ɾ������"
%>
<script language="JavaScript1.2">
<!--
function confirm_delete(){
	if (confirm("��ȷ��Ҫɾ������������")){
		return true;
	}
	return false;
}
//-->
</script>
<BR><BR><BR>
<form method="post" action="guest_mdel.asp">
<input type="hidden" name="from" value="Delete">
<input type="hidden" name="PageNo" value="<%=Request.QueryString("PageNo")%>">
<input type="hidden" name="ID" value="<%=Request.QueryString("ID")%>">
<table border=0 cellpadding=0 cellspacing=0 align=center width="506">
     <tbody> 
	   <tr><td height=20 width="504"></td></tr>
	   <tr align=center>
	     <td align=center width=504><font size="4">ɾ �� �� ��</font></td> 
	   </tr> 
      <tr><td height=20 width="504"></td></tr> 
      <tr><td height=20 width="504"></td></tr> 
	   <tr> 
	    <td align=center width="504"><img src="./images/passwd.gif" align=absmiddle>  
          ��������:&nbsp;<input type="password" style="background-color: #FFFFFF; color: #000000; border: 1 solid #000000"  name="password" size="24" maxlength="10" <%=request.cookies("UserInfo")("guestPassWord")%>"></td>  
      </tr>  
      <tr><td height=20 width="504"></td></tr>  
	  <tr>  
	   <td align=center width="504">  
       <input type="submit" style="font-size: 10pt; background-color: #C0C0C0; color: #FFFFFF" value="  ȷ   ��  " onclick="return confirm_delete()">&nbsp;&nbsp;   
       <input type="button" style="font-size: 10pt; background-color: #C0C0C0; color: #FFFFFF" value="  ȡ   ��  " onclick="history.go(-1)"></td></tr>   
      <tr><td height=20 width="504"></td></tr>   
       </form></table>   <BR><BR>
<%Bottom%>
