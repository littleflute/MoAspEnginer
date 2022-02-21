<!--#include file="rscoon.asp"-->
<% toptitle="查看所有统计" %>
<!--#include file="mzconst.asp"-->
<!--#include file="mztop.asp"-->
  <script>
 function GoToWhere(s)
 {
 var d = s.options[s.selectedIndex].value;
 window.location=d;
 s.selectedIndex=0;
 }
</script>
<div align="center">
  <table width="766" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td height="18">*--插图</td>
  </tr>
</table>
  <table width="766" border="0" align="center" cellpadding="0" cellspacing="0">
    <%   dim currentpage,id,page_total,i
           set rs = server.createobject("adodb.recordset")
		   sql="select tjid,tjtitle from mztj order by tjtime desc"
  			rs.open sql,conn,1,1
			if not rs.eof then
			 if not isempty(request.querystring("page")) then
      		currentPage=cint(request.querystring("page"))
         	else
      		currentPage=1
        	end if  
              rs.pagesize=maxsize
			  if currentpage>rs.pagecount then
			  currentpage=rs.pagecount
			  response.write"<script>alert('请不要在地址栏随意的输入页码');</script>"
			  end if
              rs.absolutepage=currentPage
              page_total=rs.pagecount 
			  do while not rs.eof and maxsize>0 
                maxsize=maxsize-1
				  %>
  <tr <% if maxsize mod 2 = 0 then%>bgcolor="#FFFFFF"<%else%>bgcolor="#D6E8E9"<%end if %>> 
    <td width="4%"> <div align="center"><img src="images/rj-tp3.gif" width="10" height="10" align="absmiddle"></div></td>
    <td width="96%" height="20"><a href="listpy.asp?id=<%=rs("tjid")%>"><%=rs("tjtitle")%></a></td>
  </tr>
  <%  rs.movenext
       loop
	   else 
	   response.write"<tr><td>对不起暂无新闻</td></tr>"
        end if
  response.write"<tr bgcolor=#F7F7F7><td height=20 colspan=2>"
      if page_total>1 then
	       response.write"总共有:["&page_total&"]页.请选择页码:"
		   if currentPage=1 then
	       response.write"上一页"
	       else 
      response.write"<a href='listalltj.asp?page="&currentPage-1&"'>上一页</a>" 
      end if  
 response.write"<select  onChange=GoToWhere(this) name=select5 class=new>"
	      for i=1 to page_total
	        if i=currentPage then 
        response.write"<option value='listalltj.asp?page="&i&"' selected>"&i&"</option>"
         else 
        response.write"<option value='listalltj.asp?page="&i&"'>"&i&"</option>"
        end if 
			next 
     response.write"</select>" 
       if currentPage=page_total then
		response.write"下一页"
		else 
      response.write"<a href='listalltj.asp?page="&currentPage+1&"'>下一页</a>" 
       end if
	  else response.write"<div align='center'>■----此版面暂时只有1页---■</div>"
		  end if 
    response.write"</td></tr><tr bgcolor=#FFFFFF><td height=20 colspan=2> <div align=right><a href='javascript:window.close()'>【关闭窗口】</a></div></td></tr></table></div>"
 call mzfoot()
%>
