<!--#include file="../mzfunc.asp"-->
<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
  <% dim vdtitle,introvd,filename,errorcount
       errorcount=""
       vdtitle=trim(request.Form("vdtitle"))
	   introvd=trim(request.Form("introvd"))
	   filename=request.Form("filename")
	   if filename="" or isnull(filename) then
	      errorcount="����"
	   response.write"<script>alert('���ϴ���Ƶ�ļ���');history.back();</script>"
            end if
			if vdtitle="" or introvd="" then
			response.write"<script>alert('���������д��ϸ');</script>"
			errorcount=delbackdb(filename)
			response.write"<script>alert('"&errorcount&"');history.back();</script>"
			  end if
			if errorcount="" then
			sql="insert into video(vdtitle,introvd,movies,times) values('"&vdtitle&"','"&introvd&"','"&filename&"','"&now()&"')"
			    conn.execute(sql)
			    end if
			%>
<div align="center">
  <table width="400" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" bgcolor="#F0F0F0">
    <tr>
      <td height="36" bgcolor="#B1E1F3"> 
        <div align="center"><strong>�ļ��ϴ��ɹ�</strong></div></td>
  </tr>
  <tr>
    <td><div align="center">���´��ϴ��ļ�ʱһ��Ҫ���������������д��ϸ,���ⷢ������Ҫ�Ĵ���</div></td>
  </tr>
  <tr>
    <td><div align="center"><a href="javascript:window.close()">���رմ��ڡ�</a></div></td>
  </tr>
</table>
</div>
<!--#include file="adminfoot.asp"-->
</body>
</html>
