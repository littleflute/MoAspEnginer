<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
    dim id,action,errorcount
	errorcount=0
	id=trim(request.QueryString("id"))
	action=trim(request.querystring("action"))
	if id="" or action="" then
	response.write"<script>alert('对不起,操作中数据丢失\n请返回重试');history.back();</script>"
       errorcount=1
	end if
	if action="del" and  errorcount=0 then
	sql="delete * from votes where voteid="&id
	conn.execute(sql)
%>
<table width="100%" border="0" align="center">
  <tr> 
    <th height="30" bgcolor="#B1E1F3"> <div align="center"><strong>删除成功</strong></div></th>
  </tr>
  <tr> 
    <td>文件已经删除,删除的数据将永久丢失,操作时请小心处理.请返回上一页面后刷新页面进行查看</td>
  </tr>
</table>
<% end if
  if action="edit" then
    sql="select voteid,votetitle from votes where voteid="&id
	set rs=conn.execute(sql)
	if not rs.eof then
  %>
  <form name="form1" method="post" action="delvotes.asp?id=<%=rs("voteid")%>&action=xiugai">
<table width="350" border="0" align="center">
    <tr bgcolor="#B1E1F3"> 
      <th height="30" colspan="2"> <div align="center"><strong>编辑投票选项</strong></div></th>
  </tr>
  <tr> 
    <td width="74">修改:</td>
    <td width="266">
        <input type="text" name="votetitle" value="<%=rs("votetitle")%>">
      </td>
  </tr>
  <tr> 
    <td colspan="2"><div align="center">
          <input type="submit" name="Submi" value="修改">
        </div></td>
  </tr>
</table>
</form>
<% else
response.write"<script>alert('你所要操作的对象不存在,请返回重试');history.bakc();</script>"
  end if
  end if
  if action="xiugai" then
  dim votetitle,Submi
     votetitle=trim(request.form("votetitle"))
	 Submi=request.form("Submi")
	 if votetitle="" or Submi<>"修改" then
response.write"<script>alert('非法操作,出现未知错误00ecaX7ae501.请返回重试');history.bakc();</script>"
          else
		  sql="update votes set votetitle='"&votetitle&"' where voteid="&id
		  conn.execute(sql)
		  end if
%>
<table width="100%" border="0" align="center">
  <tr> 
    <th height="30" bgcolor="#B1E1F3"> <div align="center"><strong>修改成功</strong></div></th>
  </tr>
  <tr> 
    <td>文件已经修改,修改前的数据将永久丢失,操作时请小心处理.请返回上一页面后刷新页面进行查看</td>
  </tr>
</table>
<% end if %>
</body>
</html>
