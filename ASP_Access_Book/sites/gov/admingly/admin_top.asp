<!--#include file="admintop.asp"-->
<!--#include file="../topconst.asp"-->
<!--#include file="../mzfunc.asp"-->
<%	
     
	 dim act,objfso,objStream,strContent,inframetitle,inframebody,inputerr
	 act=request.querystring("act")
	 if act="save" then
	 inframetitle=trim(request.form("frametitle"))
	 inframebody=server.HTMLEncode(trim(request.Form("body")))
	 if inframebody="" or inframetitle="" then
	 response.write"<script>alert('�Բ���,��Ѳ�����д��');history.back();</script>"
	           else
			end if
	 strcontent="<%"&vbcrlf
	 strcontent=strcontent&" dim frametitle "&vbcrlf
	 strcontent=strcontent&" dim framebody "&vbcrlf
	 strcontent=strcontent&" dim frametime "&vbcrlf
	 strcontent=strcontent&" frametitle="&chr(34)&inframetitle&chr(34)&vbcrlf
	 strcontent=strcontent&" framebody="&chr(34)&inframebody&chr(34)&vbcrlf
	 strcontent=strcontent&" frametime=#"&now()&"#"&vbcrlf
        strcontent=strcontent&"%"&">"&vbcrlf
		set objFso=Server.CreateObject("Scripting.FileSystemObject")
		set objStream=objFso.OpenTextFile(Server.MapPath("../topconst.asp"),2,true)
		objStream.write(strContent)
		objStream.close
		set objStream=nothing
		set objFso=nothing
	Response.Write "<script>alert('�������óɹ�,��������');window.location.href='adminbody.asp';</script>"
        	end if
     if act="" or isnull(act) then
	 if comp_check("Scripting.FileSystemObject") then
	 %>
	 <!--#include file="code.asp"-->
<form name="myform" method="post" action="admin_top.asp?act=save">
  <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr> 
      <td height="30" colspan="2" bgcolor="#B1E1F3"> <div align="center"><strong>�����������</strong></div></td>
    </tr>
    <tr> 
      <td>����:</td>
      <td><input name="frametitle" type="text" id="frametitle" value="<%=frametitle%>"></td>
    </tr>
    <tr bgcolor="#B1E1F3"> 
      <td height="25" colspan="2"> <div align="center"><strong>��������</strong></div></td>
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
          <textarea name="body" cols="70" rows="15" id="body"><%=framebody%></textarea>
        </div></td>
    </tr>
    <tr> 
      <td colspan="2"> <div align="center"> 
          <input type="submit" name="ccc" value="�� ��">
          &nbsp;&nbsp;&nbsp;&nbsp;
          <input type="reset" name="Submit2" value="ȡ ��">
        </div></td>
    </tr>
  </table>
</form>
<% else response.write"<script>alert('�Բ���,�˷�������֧���ļ���д');window.loaction.href='adminbody.asp';</script>" 
  end if
  end if %>
</body>
</html>

