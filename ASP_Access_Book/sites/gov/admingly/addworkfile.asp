<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<% 
   dim id,action
   id=request.querystring("id")
   action=request.QueryString("action")
   if action="" or isnull(action) then
%>
<form name="frmCtoy" method="post" action="addworkfile.asp?action=save&id=<%=id%>">
  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000000">
    <tr bgcolor="#B4E7F1"> 
      <td height="30" colspan="2"> 
        <div align="center"><strong>��Ӹ���</strong></div></td>
    </tr>
    <tr> 
      <td width="22%">��������:</td>
      <td width="78%" height="25"> <input name="workfile" type="text" size="30"></td>
    <tr> 
      <td height="40">�ϴ�������</td>
      <td height="25"> <iframe frameborder="0" width="400" height=24 scrolling="no" src="upload1.asp" ></iframe> 
        <input type="hidden" name="filename"> </td>
    </tr>
    <tr> 
      <td height="25">&nbsp;</td>
      <td><input type="submit" name="ups" value="ȷ��"></td>
    </tr>
  </table>
</form>
<%
     end if
	 if action="save" then
	   dim errorcount
	      errorcount=0
	 workfile=request.Form("filename")
	 filename=request.Form("workfile")
	 if workfile="" or instr(workfile,".")<=0 then
	 response.Write"<script>alert('���ϴ��ļ�');history.back();</script>"
	    errorcount=1
		 else
		 if filename="" then
     response.Write"<script>alert('����д��������');history.back();</script>"
	       errorcount=1
		   end if
		   end if
	   if errorcount=0 then
	 sql="select filename,workfile from worknet where id="&id
	  Set rs = Server.CreateObject("ADODB.RecordSet")
        rs.open sql,conn,1,3
	          end if
	  if not rs.eof then
         if rs("filename")="" then
		    rs("filename")=filename	     
            rs("workfile")=workfile
			else
			rs("filename")=rs("filename")&";"&filename
			 rs("workfile")=rs("workfile")&";"&workfile
			 end if 
			 rs.update
			 %>
			 
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td height="25" bgcolor="#B4E7F1"> 
      <div align="center"><strong>������</strong></div></td>
  </tr>
  <tr>
    <td>����Է�����һҳ����в鿴</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_worknet.asp">�����ء�</a></div></td>
  </tr>
</table>

			 <% else
			 response.Write"<script>alert('��������ʧ\n�뷵������');history.back();</script>"
                end if
				end if
			   %>
</body>
</html>
