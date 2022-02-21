<!--#include file="rscoon.asp"-->
<% toptitle="网上办事" %>
<!--#include file="mztop.asp"-->
<!--#include file="mzfunc.asp"-->
  <%  dim id,nettitle,netbody,filename,workfile
         id=request.querystring("id")
		 if id="" or isnull(id) or isnumeric(id)<>true then
		 response.write"<script>alert('错误,请返回');history.back();</script>"
            else
			sql="select nettitle,netbody,filename,workfile from worknet where id="&id
			set rs=conn.execute(sql)
		     end if
			 if rs.eof then
			 response.write"<script>alert('对不起,您所要查询的对象不存在');history.back();</script>"	
               else
			   nettitle=rs("nettitle")
			   netbody=rs("netbody")
			   filename=rs("filename")
			   workfile=rs("workfile")
                     end if
			  %>
<table width="766" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr align="center" bgcolor="#F6F6F6"> 
    <td height="27" colspan="3"  background="images/zxzctp2.gif"><%=nettitle%></td>
  </tr>
  <tr> 
    <td height="20" background="images/dhbg2.gif" bgcolor="#F6F6F6"> 
      <div align="center"><strong>相关信息</strong></div></td>
    <td colspan="2">*--插图</td>
  </tr>
  <tr> 
    <td width="16%" rowspan="2" valign="center" bgcolor="#F6F6F6"> <div align="center"><br>
        <%   dim pytitle
	   sql="select top 10 id,nettitle from worknet where id<>"&id&" order by nettime desc"
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
        <a href="worknet.asp?id=<%=pytitle(0,i)%>"><%=pytitle(1,i)%></a><br>
        <% next 
	    else
		response.write"暂无其它文章"
		end if
	   %>
      </div></td>
    <td width="1%" rowspan="2" valign="top"><img src="images/dhtp2.gif" width="2" height="101%"></td>
    <td width="83%" height="80%" valign="top">&nbsp;&nbsp;<%=ubbcode(netbody)%></td>
  </tr>
  <tr>
    <td height="20%"> 
      <%
	  if filename<>"" and isnull(filename)<>true then 
	filename=split(filename,";")
	workfile=split(workfile,";")
	for i=0 to ubound(filename)
response.write "<a href='download.asp?n=admingly/movie/"&workfile(i)&"'>"&filename(i)&"[点击下载]</a><br>"
	 next 
	 end if
 response.write"</td></tr><tr align='right' bgcolor=#F6F6F6>"&_
 "<td height='20' colspan='3'> <a href='javascript:window.close()'>【关闭窗口】</a></td></tr></table>"
 call mzfoot() %>
