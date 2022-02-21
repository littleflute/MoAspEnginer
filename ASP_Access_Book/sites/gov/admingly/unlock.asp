<!--#include file="admintop.asp"-->
<!--#include file="../mzconst.asp"-->
<!--#include file="isadmin.asp"-->
<form name="form1" method="POST">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr bgcolor="#B1E1F3"> 
      <td height="30" colspan="5"> 
        <div align="center"><strong>所有被锁定的IP</strong></div></td>
    </tr>
    <tr> 
      <td width="33%" align="center">ip地址</td>
      <td width="24%" align="center">被锁定的时间</td>
      <td width="21%" align="center"> 使用账号</td>
      <td width="10%" align="center">操作</td>
      <td width="12%" align="center">选择</td>
    </tr>
    <%   dim currentpage,id,page_total,i 
   sql="select * from killip order by times desc"
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
      <td width="33%" align="center"><%=rs("badip")%></td>
      <td width="24%" align="center"><%=rs("times")%></td>
      <td width="21%" align="center"><%=rs("username")%></td>
      <td width="10%"><div align="center"><a href="dellockip.asp?id=<%=rs("id")%>">解除</a></div></td>
      <td width="12%"><div align="center"> 
          <input type="checkbox" name="Board" id="<%=rs("id")%>" style="font-size: 9pt">
        </div></td>
    </tr>
    <%  rs.movenext
      loop
	  else response.write"<tr><td colspan='4'>暂无ip被锁</td></tr>"
	  end if
  %>
    <tr> 
      <td colspan="5"> <div align="center"> &nbsp;&nbsp; <a href="addlockip.asp">添加ip地址</a>&nbsp;&nbsp; 
          <input type="button" value="全 选" onclick="sltAll()" name=button1>
          &nbsp;&nbsp; 
          <input type="button" value="清 空" onclick="sltNull()" name=button2>
          &nbsp;&nbsp; 
          <input type="submit" value="解 除" name="tijiao" onclick="SelectChk('dellockip.asp')">
        </div></td>
    </tr>
    <tr> 
      <td colspan="5"> <%if page_total>1 then
	       response.write"总共有:"
	       response.Write(page_total)
	       response.write"页."
	       response.write "请选择页码:"
		   if currentPage=1 then
	       response.write"上一页"
	       else %> <a href="unlock.asp?page=<%=currentPage-1%>">上一页</a> 
        <%
	                   end if  %> <script>
 function GoToWhere(s)
 {
 var d = s.options[s.selectedIndex].value;
 window.location=d;
 s.selectedIndex=0;
 }
//-->
</script> <select  onChange=GoToWhere(this) name=select5 class="new">
          <%  
	      for i=1 to page_total
	        if i=currentPage then 
		  %>
          <option value="unlock.asp?page=<%=i%>" selected><%=i%></option>
          <% else %>
          <option value="unlock.asp?page=<%=i%>"><%=i%></option>
          <%  end if 
			next %>
        </select> <% if currentPage=page_total then
		response.write"下一页"
		else %> <a href="unlock.asp?page=<%=currentPage+1%>">下一页</a> 
        <% end if
		 else response.write"所有内容只有一页,当前页面:第1页"
		  end if %></td>
    </tr>
  </table>
</form>
</body>
</html>

