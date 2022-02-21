<!--#include file="admintop.asp"-->
<!--#include file="../mzconst.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%	
     dim act,objfso,objStream,strContent,inputmax,inputtop,inputerr
	 dim inputiscook
	 act=request.querystring("act")
	 if act="save" then
	 inputiscook=request.form("iscok")
	 inputmax=trim(request.form("maxsize"))
	 inputerr=trim(request.form("errorpwd"))
	 inputtop=server.HTMLEncode(trim(request.Form("toppic")))
	 if inputmax="" or inputerr="" or inputtop="" or inputiscook="" then
	 response.write"<script>alert('对不起,请把参数填写完');history.back();</script>"
	           else
	 if isnumeric(inputmax)<>true or isnumeric(inputerr)<>true then
	 response.write"<script>alert('有些参数只能是数字不能有字符');history.back();</script>"	   
	        end if
			end if
	 strcontent="<%"&vbcrlf
	 strcontent=strcontent&" dim maxsize "&vbcrlf
	 strcontent=strcontent&" dim errorpwd "&vbcrlf
	 strcontent=strcontent&" maxsize="&inputmax&vbcrlf
	 strcontent=strcontent&" errorpwd="&inputerr&vbcrlf
	 strcontent=strcontent&"const iscookies="&inputiscook&vbcrlf
	 strcontent=strcontent&"const toppic="&chr(34)& inputtop &chr(34)&vbcrlf
        strcontent=strcontent&"%"&">"&vbcrlf
		set objFso=Server.CreateObject("Scripting.FileSystemObject")
		set objStream=objFso.OpenTextFile(Server.MapPath("../mzconst.asp"),2,true)
		objStream.write(strContent)
		objStream.close
		set objStream=nothing
		set objFso=nothing
	Response.Write "<script>alert('保存设置成功,即将返回');window.location.href='adminbody.asp';</script>"
        	end if
     if act="" or isnull(act) then
	 if comp_check("Scripting.FileSystemObject") then
	 %>
  <form name="form1" method="post" action="admin_basic.asp?act=save">
  <table width="80%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr bgcolor="#B1E1F3"> 
      <td height="30" colspan="2"> 
        <div align="center"><strong>基本参数设定</strong></div></td>
    </tr>
    <tr> 
      <td>每页显示新闻条数:</td>
      <td> <input type="text" name="maxsize" value=<%=maxsize%>> 
      </td>
    </tr>
	<tr><td>允许输错几次密码:</td><td><input type="text" name="errorpwd" value=<%=errorpwd%>></td></tr>
    <tr>
      <td>是否允许cookies:</td>
      <td>允许: 
        <input name="iscok" type="radio" value="true" checked>
        　禁止: 
        <input type="radio" name="iscok" value="false">
      </td>
    </tr>
	<tr> 
      <td>主页面滚动标语:</td>
      <td><textarea name="toppic" cols="40" rows="10"><%=toppic%></textarea></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center"> 
          <input type="submit" name="Submit" value="确 定">
          　
<input type="submit" name="Submit2" value="取 消">   </div></td>
    </tr>
  </table>
</form>
<% else response.write"<script>alert('对不起,此服务器不支持文件读写');window.loaction.href='adminbody.asp';</script>" 
  end if
  end if %>
</body>
</html>
