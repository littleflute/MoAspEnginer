<!--#include file="admintop.asp"-->
<!--#include file="../mzconst.asp"-->
<!--#include file="isadmin.asp"-->
<%
  dim id
  id=request.QueryString("id")
    if id="" or isnull(id) or isnumeric(id)<>true then
	response.write"<script>alert('ĪҪ����\nĪҪ����');window.location.href='error.htm';</script>"
	         else
			 sql="update admingly set islock="&errorpwd&" where id="&id
			 conn.execute(sql)
			response.write"<table width=350 border=1 align='center' cellpadding='0' cellspacing='0' bordercolor='#000000'>"&_
			"<tr><td height=30 align=center bgcolor=#B1E1F3><strong>�����ɹ�</strong></td></tr>"&_
			"<tr><td align=center>�뷵����һҳ����в鿴</td></tr><tr><td align=center><a href='listalladmin.asp'>�����ء�</a></td></tr></table>"
			 end if
	        response.write"</body></html>"
			%>

  
  
