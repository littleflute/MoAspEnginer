<!--#include file="../mzfunc.asp"-->
<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
  <% dim pytitle,pybody,filename,errorcount
       errorcount=""
       pytitle=trim(request.Form("pytitle"))
	   pybody=server.HTMLEncode(trim(request.Form("pybody")))
	   filename=request.Form("filename")
			if pytitle="" or pybody="" then
			response.write"<script>alert('���������д��ϸ');</script>"
			  if trim(filename)<>"" then
			errorcount=delbackdb(filename)
			response.write"<script>alert('"&errorcount&"');history.back();</script>"
			   end if
			  end if
			if errorcount="" then
			sql="insert into policy(pytitle,pybody,pyimg,times) values('"&pytitle&"','"&pybody&"','"&filename&"','"&now()&"')"
			    conn.execute(sql)
			    end if
			%>
<div align="center">
  <table width="400" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" bgcolor="#F0F0F0">
    <tr>
      <td height="36" bgcolor="#B1E1F3"> 
        <div align="center"><strong>��ӳɹ�</strong></div></td>
  </tr>
  <tr>
    <td><div align="center">�ϴ��ļ�ʱһ��Ҫ���������������д��ϸ,���ⷢ������Ҫ�Ĵ���</div></td>
  </tr>
  <tr>
    <td><div align="center"><a href="javascript:window.close()">���رմ��ڡ�</a></div></td>
  </tr>
</table>
</div>
<!--#include file="adminfoot.asp"-->
</body>
</html>
