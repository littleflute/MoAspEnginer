<!--#include file="admintop.asp"-->
<!--#include file="code.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%  dim action
    action=request.querystring("action")
	if action="" then
%>
<form name="myform" method="post" action="addtj.asp?action=add">
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr bgcolor="#B1E1F3"> 
      <td height="30" colspan="2"> 
        <div align="center"><strong>���ͳ��</strong></div></td>
    </tr>
    <tr> 
      <td width="17%">ͳ�Ʊ���:</td>
      <td width="83%"><input name="tjtitle" type="text" id="title"></td>
    </tr>
    <tr bgcolor="#B1E1F3"> 
      <td height="25" colspan="2"> <div align="center"><strong>ͳ������</strong></div></td>
    </tr>
    <tr>
      <td colspan="2"><div align="center"><img onClick=bold() src="pic/icon_editor_bold.gif" width="22" height="22" alt="����" border="0"> 
          <img onClick=italicize() src="pic/icon_editor_italicize.gif" width="23" height="22" alt="б��" border="0"> 
          <img onClick=underline() src="pic/icon_editor_underline.gif" width="23" height="22" alt="�»���" border="0"> 
          <img onClick=center() src="pic/icon_editor_center.gif" width="22" height="22" alt="����" border="0"> 
          <img onClick=hyperlink() src="pic/icon_editor_url.gif" width="22" height="22" alt="��������" border="0"> 
          <img onClick=email() src="pic/icon_editor_email.gif" width="23" height="22" alt="Email����" border="0"> 
          <img onClick=image() src="pic/icon_editor_image.gif" width="23" height="22" alt="ͼƬ" border="0"> 
          <img onClick=flash() src="pic/swf.gif" width="23" height="22" alt="FlashͼƬ" border="0"> 
          <img onClick=showcode() src="pic/icon_editor_code.gif" width="22" height="22" alt="���" border="0"> 
          <img onClick=quote() src="pic/icon_editor_quote.gif" width="23" height="22" alt="����" border="0"> 
          <img onClick=list() src="pic/icon_editor_list.gif" width="23" height="22" alt="Ŀ¼" border="0"> 
          <img onClick=setfly() height=22 alt=������ src="pic/fly.gif" width=23 border=0> 
          <img onClick=move() height=22 alt=�ƶ��� src="pic/move.gif" width=23 border=0> 
          <img onClick=glow() height=22 alt=������ src="pic/glow.gif" width=23 border=0> 
          <img onClick=shadow() height=22 alt=��Ӱ�� src="pic/shadow.gif" width=23 border=0></div></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center"> 
          <textarea name="body" cols="70" rows="15" id="body"></textarea>
        </div></td>
    </tr>
    <tr> 
      <td colspan="2"> <div align="center"> 
          <input type="submit" name="ccc" value="���">
          &nbsp;&nbsp;&nbsp;&nbsp;
          <input type="reset" name="Submit2" value="ȡ��">
        </div></td>
    </tr>
  </table>
</form>
<%  end if
       if action="add" then
	   dim tjtitle,body
	   tjtitle=trim(request.form("tjtitle"))
	   body=trim(request.Form("body"))
	   if tjtitle="" or body="" then
	   response.write"<script>alert('���������д��ϸ\n�뷵������');history.back();</script>"
             else
		  body=server.htmlencode(body)
		  sql="insert into mztj(tjtitle,tjbody,tjtime) values('"&tjtitle&"','"&body&"','"&now()&"')"
		  conn.execute(sql)
		     end if
		%>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <th height="30" bgcolor="#B1E1F3">��ӳɹ�</th>
  </tr>
  <tr>
    <td>�����ѳɹ����,��ѡ����Ҫ���ص�ҳ��</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_tj.asp">��������һҳ�桿</a><a href="adminbody.asp">�����ع�����ҳ��</a></div></td>
  </tr>
</table>
<% end if %>
</body>
</html>
