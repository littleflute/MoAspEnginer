<!--#include file="admintop.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%  dim id,movies,vdtitle,introvd
   id=request.querystring("id")
   action=request.querystring("action")
   if id="" or isnull(id) then
   response.write"<script>alert('对不起,操作出现错误');history.back();</script>"
    else
	sql="select * from video where id="&id
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open sql,conn,1,3
	   end if
       if not rs.eof then 
	   if action="" then%>
<form name="frmCtoy" method="post" action="editvideo.asp?action=edit&id=<%=rs("id")%>" >
  <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000" bordercolorlight="#000000">
    <tr bgcolor="#B1E1F3"> 
      <td height="30" colspan="2"> 
        <div align="center"><strong>修改视频文件</strong></div></td>
  </tr>
  <tr> 
    <td width="35%">视频文件标题:</td>
    <td height="25"> <input name="vdtitle" type="text" size="30" value="<%=rs("vdtitle")%>"></td>
  <tr> 
  <tr> 
    <td width="35%">内容简介:</td>
    <td width="65%"> <textarea name="introvd" cols="30" rows="5"><%=rs("introvd")%></textarea> </td>
  </tr>
  <tr> 
    <td height="27">视频文件名：</td>
    <td height="25"><input name="movies" type="text" size="30" value="<%=rs("movies")%>"></td>
  </tr>
  <tr> 
    <td height="25">&nbsp;</td>
    <td><input type="submit" name="ups" value="确定"></td>
  </tr>
</table>
</form>
<%  else
     if action="edit" then
	 vdtitle=request.form("vdtitle")
	 introvd=request.form("introvd")
	 movies=request.form("movies")
	 if vdtitle<>"" and introvd<>"" and movies<>"" then
	    if not isexists(movies) then
		response.write"<script>alert('对不起,你所填的文件并不存在\n请返回重试');history.back();</script>"
	        else
	 rs("vdtitle")=vdtitle
	 rs("introvd")=introvd
	 rs("movies")=movies
	 rs.update 
	 end if%>
	
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td height="30" bgcolor="#B1E1F3"> 
      <div align="center"><strong>修改成功</strong></div></td>
  </tr>
  <tr>
    <td>数据修改成功,你所修改后的数据内容将直接在主界在上显示.操作时请谨慎处理</td>
  </tr>
  <tr>
    <td height="41">
<div align="center"><a href="video.asp">【返回】</a><a href="adminbody.asp">【继续其它操作】</a></div></td>
  </tr>
</table>
 
<%	 else response.write"<script>alert('对不起,请把资料填写详细');history.back();</script>"
       end if
	   end if
	   end if
	   else  response.write"<script>alert('对不起,你所要操作的对象未找到.');history.back();</script>"
        end if
%>
</body>
</html>
