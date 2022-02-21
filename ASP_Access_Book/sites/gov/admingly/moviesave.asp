<!--#include file="../mzfunc.asp"-->
<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
  <% dim vdtitle,introvd,filename,errorcount
       errorcount=""
       vdtitle=trim(request.Form("vdtitle"))
	   introvd=trim(request.Form("introvd"))
	   filename=request.Form("filename")
	   if filename="" or isnull(filename) then
	      errorcount="错误"
	   response.write"<script>alert('请上传视频文件啊');history.back();</script>"
            end if
			if vdtitle="" or introvd="" then
			response.write"<script>alert('请把资料填写详细');</script>"
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
        <div align="center"><strong>文件上传成功</strong></div></td>
  </tr>
  <tr>
    <td><div align="center">请下次上传文件时一定要把所在填定的内容填写详细,以免发生不必要的错误</div></td>
  </tr>
  <tr>
    <td><div align="center"><a href="javascript:window.close()">【关闭窗口】</a></div></td>
  </tr>
</table>
</div>
<!--#include file="adminfoot.asp"-->
</body>
</html>
