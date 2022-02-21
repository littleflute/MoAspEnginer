<!--#include file="rscoon.asp"-->
<% toptitle="关注政策法规" %>
<!--#include file="mztop.asp"-->
<!--#include file="mzfunc.asp"-->
<table width="766" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr> 
    <%  dim pyid
         pyid=request.querystring("id")
		 if pyid="" or isnull(pyid) or isnumeric(pyid)<>true then
		 response.write"<script>alert('错误,请返回');history.back();</script>"
            else
			sql="select pytitle,pybody,pyimg from policy where id="&pyid
			set rs=conn.execute(sql)
		     end if
			 if rs.eof then
			 response.write"<script>alert('对不起,您所要查询的对象不存在');history.back();</script>"	
               else
			   pytitle=rs("pytitle")
			   pybody=rs("pybody")
			   pyimg=rs("pyimg")
                     end if
			  %>
    <td height="25" colspan="3" background="images/zxzctp2.gif" bgcolor="#F6F6F6"> 
      <div align="center"><%=pytitle%></div></td>
  </tr>
  <tr> 
    <td height="30" background="images/dhbg2.gif" bgcolor="#F6F6F6"> <div align="center"><strong>相关信息</strong></div></td>
    <td colspan="2">*--插图</td>
  </tr>
  <tr> 
    <td width="16%" valign="center" bgcolor="#F6F6F6"> <div align="center"><br>
        <%   dim pytitle
	   sql="select top 10 id,pytitle from policy where id<>"&pyid&" order by times desc"
	       set rs=conn.execute(sql)
		   if not rs.eof then
		   pytitle=rs.getrows
		    if ubound(pytitle,2)<9 then
			   maxcount=ubound(pytitle,2)
				 else
			   maxcount=9
				 end if
			  for i=0 to maxcount
	   %>
        <a href="listpy.asp?id=<%=pytitle(0,i)%>"><%=pytitle(1,i)%></a><br>
        <% next 
	    else
		response.write"暂无其它文章"
		end if
	   %>
      </div></td>
    <td width="1%" valign="top"><a href="listpy.asp?id=<%=pyid%>"></a><img src="images/dhtp2.gif" width="2" height="100%"></td>
    <td width="83%" valign="top"><% if pyimg<>"" then %> <div align="center"><IMG height=100  src="admingly/movie/<%=pyimg%>" width=100></div>
      <br> <% end if %> &nbsp;&nbsp;<%=ubbcode(pybody)%></td>
  </tr>
  <tr> 
    <td height="20" colspan="3" bgcolor="#F6F6F6"> <div align="right"><a href="javascript:window.close()">【关闭窗口】</a></div></td>
  </tr>
</table>
<% call mzfoot() %>
