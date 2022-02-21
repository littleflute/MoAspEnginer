<!--#include file="admintop.asp"-->
<!--#include file="../mzconst.asp"-->
<!--#include file="isadmin.asp"-->
<form name="form1" method="POST">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr bgcolor="#B1E1F3"> 
      <td height="30" colspan="5"> <div align="center"><strong>所有投票统计</strong></div></td>
    </tr>
    <tr> 
      <td width="45%"><div align="center">投票标题</div></td>
      <td width="18%"><div align="center">状态</div></td>
      <td colspan="2"><div align="center">操作</div></td>
      <td width="13%"><div align="center">选择</div></td>
    </tr>
    <%   dim currentpage,id,page_total,i 
   sql="select id, isvote, title from stat order by id desc"
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
      <td width="45%"><div align="center"><a href="listvote.asp?id=<%=rs("id")%>"><%=rs("title")%></a></div></td>
      <td width="18%"><div align="center">
          <% if trim(rs("isvote"))="true" then response.write"开启状态" else response.write"关闭状态" end if%>
        </div></td>
      <td width="17%"><div align="center"><a href="editvote.asp?id=<%=rs("id")%>">
          修改</a>/<a href="delstat.asp?id=<%=rs("id")%>">删除</a></div></td>
      <td width="7%"><% dim on_off
	  if trim(rs("isvote"))="true" then on_off="关闭" else on_off="开启"  end if%>
	  <input type="button" name="on_off" value="<%=on_off%>" onclick="BoardWin('on_off.asp?id=<%=rs("id")%>')"></td>
      <td width="13%"><div align="center"> 
          <input type="checkbox" name="Board" id="<%=rs("id")%>" style="font-size: 9pt">
        </div></td>
    </tr>
    <%  rs.movenext
      loop
	  else response.write"<tr><td colspan='4'>暂无内容</td></tr>"
	  end if
  %>
    <tr> 
      <td colspan="5"> <div align="center"> &nbsp;&nbsp; 
          <input type="button" value="添加投票" onclick="BoardWin('addstat.asp')" name=add>
          &nbsp;&nbsp; 
          <input type="button" value="全 选" onclick="sltAll()" name=button1>
          &nbsp;&nbsp; 
          <input type="button" value="清 空" onclick="sltNull()" name=button2>
          &nbsp;&nbsp; 
          <input type="submit" value="删 除" name="tijiao" onclick="SelectChk('delstat.asp')">
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
	       else %> <a href="admin_vote.asp?page=<%=currentPage-1%>">上一页</a> 
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
          <option value="admin_vote.asp?page=<%=i%>" selected><%=i%></option>
          <% else %>
          <option value="admin_vote.asp?page=<%=i%>"><%=i%></option>
          <%  end if 
			next %>
        </select> <% if currentPage=page_total then
		response.write"下一页"
		else %> <a href="admin_vote.asp?page=<%=currentPage+1%>">下一页</a> 
        <% end if
		 else response.write"总投票选顶只有一页,当前页面:第1页"
		  end if %></td>
    </tr>
  </table>
</form>
</body>
</html>
