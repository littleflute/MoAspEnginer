<!--#include file="admintop.asp"-->
<!--#include file="code.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%
   dim id,action
   id=request.querystring("id")
   action=request.querystring("action")
   if id="" or isnull("id") then
   response.write"<script>alert('�Բ���,����ʱ���ݶ�ʧ,��ҳ�治֧��ˢ��');history.back();</script>"
      end if
	  if action="" or isnull(action) then
	  sql="select id,nstitle,nsbody,writers from news where id="&id
	  set rs=conn.execute(sql)
	  if not rs.eof then
%>
<form name="myform" method="post" action="editmznew.asp?action=edit&id=<%=id%>">
  <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr> 
      <td height="30" colspan="2" bgcolor="#B1E1F3"> <div align="center"><strong>�����Ŀ����</strong></div></td>
    </tr>
    <tr> 
      <td width="20%">��Ŀ���ű���:</td>
      <td width="80%"><input name="title" type="text" id="title" value="<%=rs("nstitle")%>"></td>
    </tr>
	<tr><td>������λ:</td><td><input name="writer" type="text" value="<%=rs("writers")%>"></td></tr>
    <tr bgcolor="#B1E1F3"> 
      <td height="25" colspan="2"> 
        <div align="center"><strong>��������</strong></div></td>
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
          <textarea name="body" cols="70" rows="15" id="body"><%=rs("nsbody")%></textarea>
        </div></td>
    </tr>
    <tr> 
      <td colspan="2"> <div align="center"> 
          <input type="submit" name="ccc" value="�޸�">
          &nbsp;&nbsp;&nbsp;&nbsp;
          <input type="reset" name="Submit2" value="ȡ��">
        </div></td>
    </tr>
  </table>
</form>
<%
   else     response.write"<script>alert('�Բ���,��Ҫ����������δ�ҵ�\n�뷵������');history.back();</script>"
          end if
		  end if
		  if action="edit" then
		  dim errorcount,nstitle,writer,nsbody,ccc
		      errorcunt=0
			  nstitle=request.form("title")
			  writer=request.form("writer")
			  nsbody=request.form("body")
			  ccc=request.form("ccc")
			  if nstitle="" or writer="" or nsbody="" then
			     response.write"<script>alert('�Բ���,�뷵�ذ�������д��ϸ');history.back();</script>"
                   errorcount=1
				   else
				   if ccc<>"�޸�" then
				      response.write"<script>alert('�Բ���,�뷵�ذ�������д��ϸ');history.back();</script>"
                      errorcount=1
					  end if
				end if
				if errorcount=0 then
				   sql="update news set nstitle='"&nstitle&"', nsbody='"&nsbody&"',times='"&now()&"', writers='"&writer&"' where id="&id	  
			           conn.execute(sql)
			        %>
					<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <th height="30" bgcolor="#B1E1F3">�޸ĳɹ�</th>
  </tr>
  <tr>
    <td>�����ѳɹ��޸�,����Է��ز鿴.</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_news.asp">�����ز鿴��</a> </div></td>
  </tr>
</table>
<% end if 
   end if %>
</body>
</html>
