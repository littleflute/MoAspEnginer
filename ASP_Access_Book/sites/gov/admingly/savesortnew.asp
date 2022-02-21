<!--#include file="../mzfunc.asp"-->
<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%  dim frominto,title,body
      if request.form("ccc")="提交" then
     frominto=request.Form("sort")
	 title=request.form("title")
	 body=server.HTMLEncode(request.form("body"))
	   if frominto="" or title="" or body="" then
	   response.write"<script>alert('请把资料填定详细');history.back();</script>"
	            else
		sql="select * from allarti where title='"&title&"' and frominto="&frominto
		    set rs=conn.execute(sql)
			if not rs.eof then
		response.write"<script>alert('此新闻标题已存在,请更换新闻标题');history.back();</script>"		
	            else
	 sql="insert into allarti(title,body,frominto,times) values('"&title&"','"&body&"','"&frominto&"','"&now()&"')"
         conn.execute(sql)
		 end if
		 end if
		 else
		 response.write"<script>alert('非法操作');history.back();</script>"
            end if
%>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td height="30" bgcolor="#B4E7F1"><div align="center"><strong>添加成功</strong></div></td>
  </tr>
  <tr>
    <td><div align="center">栏目新闻已经成功添加,你可以返回主页进行查看!</div></td>
  </tr>
  <tr>
    <td><div align="center"><a href="sort_new.asp">【返回】</a><a href="adminbody.asp">【继续其它操作】</a></div></td>
  </tr>
</table>
<!--#include file="adminfoot.asp"-->
</body>
</html>
