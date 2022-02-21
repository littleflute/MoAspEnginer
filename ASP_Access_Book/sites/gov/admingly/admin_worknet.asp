<!--#include file="admintop.asp"-->
<!--#include file="../mzconst.asp"-->
<!--#include file="isadmin.asp"-->
<form name="form1" method="POST">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr bgcolor="#B1E1F3"> 
      <td height="30" colspan="4"> 
        <div align="center"><strong>所有网上事务</strong></div></td>
  </tr>
  <tr> 
    <td width="53%"><div align="center">事务文件名</div></td>
    <td width="27%"><div align="center">发布日期</div></td>
    <td width="10%"><div align="center">操作</div></td>
    <td width="10%"><div align="center">选择</div></td>
  </tr>
  <%   dim currentpage,id,page_total,i 
   sql="select id, nettitle,nettime from worknet order by nettime desc"
      Set rs = Server.CreateObject("ADODB.RecordSet")
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
  <tr> 
    <td width="53%"><div align="center"><a href="../worknet.asp?id=<%=rs("id")%>" target="_blank"><%=rs("nettitle")%></a></div></td>
    <td width="27%"><div align="center"><%=rs("nettime")%></div></td>
    <td width="10%"><div align="center"><a href="editworknet.asp?id=<%=rs("id")%>">修改</a>/<a href="delworknet.asp?id=<%=rs("id")%>">删除</a></div></td>
    <td width="10%"><div align="center">
        <input type="checkbox" name="Board" id="<%=rs("id")%>" style="font-size: 9pt">
      </div></td>
  </tr>
  <%  rs.movenext
      loop
	  else response.write"<tr><td colspan='4'>暂无内容</td></tr>"
	  end if
  %>
  <tr>
    <td colspan="4"> <div align="center"> &nbsp;&nbsp; <a href="addworknet.asp">添加内容</a>&nbsp;&nbsp; 
          <input type="button" value="全 选" onclick="sltAll()" name=button1>
        &nbsp;&nbsp;
        <input type="button" value="清 空" onclick="sltNull()" name=button2>
        &nbsp;&nbsp;
        <input type="submit" value="删 除" name="tijiao" onclick="SelectChk('delworknet.asp')">
      </div></td></tr>
  <tr> 
    <td colspan="4">
	<%if page_total>1 then
	       response.write"总共有:"
	       response.Write(page_total)
	       response.write"页."
	       response.write "请选择页码:"
		   if currentPage=1 then
	       response.write"上一页"
	       else %> <a href="admin_worknet.asp?page=<%=currentPage-1%>">上一页</a> 
                        <%
	                   end if  %>
		   <script>
 function GoToWhere(s)
 {
 var d = s.options[s.selectedIndex].value;
 window.location=d;
 s.selectedIndex=0;
 }
//-->
</script>
<select  onChange=GoToWhere(this) name=select5 class="new">
            <%  
	      for i=1 to page_total
	        if i=currentPage then 
		  %>
            <option value="admin_worknet.asp?page=<%=i%>" selected><%=i%></option>
            <% else %>
			<option value="admin_worknet.asp?page=<%=i%>"><%=i%></option>
			<%  end if 
			next %>
          </select>
		  <% if currentPage=page_total then
		response.write"下一页"
		else %> <a href="admin_worknet.asp?page=<%=currentPage+1%>">下一页</a> 
                        <% end if
		 else response.write"总新闻数只有一页,当前页面:第1页"
		  end if %></td>
  </tr>
</table>
</form>
</body>
</html>

