<!--#include file="admintop.asp"-->
<!--#include file="ipjc.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="../mzconst.asp"-->
<!--#include file="adminmd5.asp"-->
<%  if session("username")="" then
    dim action,cookiename,cookiepwd
	     if iscookies then
	cookiename=trim(request.Cookies("mzcoks")("unm"))
	cookiepwd=trim(request.Cookies("mzcoks")("upwd"))
	      end if
	if cookiename<>"" and cookiepwd<>"" then
	sql="select * from admingly where user_name='"&cookiename&"' and password='"&md5(cookiepwd)&"'"  
	     set rs=conn.execute(sql)
	     if not rs.eof then
         session("username")=cookiename
		 session("ispass")="pass" 
           end if
	     end if
		 end if
	 if  session("username")="" or session("ispass")<>"pass" then
	 response.Write"<script>alert('�Բ������ȵ�½');window.location.href='adminlogin.asp';</script>"
                 end if
	 %>
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td height="30" bgcolor="#B6D7E7"> 
      <div align="center"><strong>��½�ɹ�</strong></div></td>
  </tr>
  <tr>
    <td background="topBar_bg.gif">�����������ڴ˶Ժ�̨���в���,�������,����(����,��Ƶ�̳�,��������,������,)�����ѡ������Ҫ����Ķ���,�Ϳɶ��������Ӧ�Ĺ���.�������ɾ������Ա,��ӵĹ���ԱҲ����ͬ����Ȩ��,�����޸ĺ����μ�����,���������볬���趨����,�˹���Ա�˺Ž�������,ͬʱip��ַҲ�ᱻ����.Ҫ������һ̨����������һ������Ա�˺Ŷ�����н���.</td>
  </tr>
  <tr>
    <td background="topBar_bg.gif"> 
      <div align="center">����ʹ���������κ����ⶼ�������ǽ�����ϵ!</div>
    </td>
  </tr>
  <tr>
    <td align="center" background="topBar_bg.gif">�����ڵ�IP��ַ��:<%=getip()%></td>
  </tr>
</table>
<!--#include file="adminfoot.asp"-->
</body>
</html>
