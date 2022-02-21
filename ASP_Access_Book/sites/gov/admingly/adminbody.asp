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
	 response.Write"<script>alert('对不起，请先登陆');window.location.href='adminlogin.asp';</script>"
                 end if
	 %>
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td height="30" bgcolor="#B6D7E7"> 
      <div align="center"><strong>登陆成功</strong></div></td>
  </tr>
  <tr>
    <td background="topBar_bg.gif">　　您可以在此对后台进行操作,可以添加,管理(新闻,视频教程,友情链接,导航条,)在左侧选择您需要管理的对象,就可对其进行相应的管理.可以添加删除管理员,添加的管理员也具有同样的权限,密码修改后请牢记密码,如果输错密码超过设定次数,此管理员账号将被锁定,同时ip地址也会被锁定.要在另外一台电脑上用另一个管理员账号对其进行解锁.</td>
  </tr>
  <tr>
    <td background="topBar_bg.gif"> 
      <div align="center">如在使用中遇到任何问题都可与我们进行联系!</div>
    </td>
  </tr>
  <tr>
    <td align="center" background="topBar_bg.gif">你现在的IP地址是:<%=getip()%></td>
  </tr>
</table>
<!--#include file="adminfoot.asp"-->
</body>
</html>
