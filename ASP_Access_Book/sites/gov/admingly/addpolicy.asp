<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<form name="frmCtoy" method="post" action="policysave.asp?action=save" >
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000000">
    <tr> 
      <td height="30" colspan="2" bgcolor="#B4E7F1"> 
        <div align="center"><strong>������߷���</strong></div></td></tr>
    <tr>
      <td width="22%">���߷������:</td>
      <td height="25"> <input name="pytitle" type="text" size="30"></td>
    <tr>
	<tr>
      <td width="22%">���߷�������:</td>
      <td width="78%"> 
        <textarea name="pybody" cols="60" rows="20"></textarea>
      </td>
    </tr>
 <tr> 
      <td height="40">�ϴ�ͼƬ��</td>
      <td height="25"> 
        <iframe frameborder="0" width="400" height=24 scrolling="no" src="upload1.asp" ></iframe>
     <input type="hidden" name="filename">
      </td>
  </tr>
  <tr>
      <td height="25">&nbsp;</td>
    <td><input type="submit" name="ups" value="ȷ��"></td>
	</tr>
</table>
</form>
</body>
</html>