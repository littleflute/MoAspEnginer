<!--#include file="admintop.asp"-->
<!--#include file="../mzconst.asp"-->
<!--#include file="isadmin.asp"-->
<SCRIPT>
function checkForm(frm)
{
	if (frm.key.value == null || frm.key.value.length == 0) {
		alert("������Ҫ�����Ĺؼ���!");
		frm.keyword.focus();
		return false;
	}

	return true;
}  
</SCRIPT>
<form name="form1" method="POST">
  <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr bgcolor="#B1E1F3"> 
    <%  dim id,sorts,i
	id=request.querystring("id")
	action=request.querystring("action") 
	sql="select * from allsort order by id"
	   set rs=conn.execute(sql)
	   sorts=rs.getrows
	%>
	<td height="30" colspan="4"> <div align="center"><strong>��ѡ����Ŀ:</strong>&nbsp;&nbsp;
	    <script>
 function GoToWhere(s)
 {
 var d = s.options[s.selectedIndex].value;
 window.location=d;
 s.selectedIndex=0;
 }
//-->
</script> <select  onChange=GoToWhere(this) name=select5 class="new">
       <option value="sort_new.asp">������Ŀ</option>
	    <%  
	      for i=0 to ubound(sorts,2)
	        if cint(sorts(0,i))=cint(id) then 
		  %>
        <option value="sort_new.asp?id=<%=sorts(0,i)%>" selected><%=sorts(1,i)%></option>
        <% else %>
        <option value="sort_new.asp?id=<%=sorts(0,i)%>"><%=sorts(1,i)%></option>
        <%  end if 
			next %>
      </select> 
	</div></td>
  </tr>
  <tr> 
    <td width="53%"><div align="center">��Ƶ�ļ�</div></td>
    <td width="27%"><div align="center">�ϴ�����</div></td>
    <td width="10%"><div align="center">����</div></td>
    <td width="10%"><div align="center">ѡ��</div></td>
  </tr>
  <%   dim currentpage,page_total
  if action="search" then
     dim keys
	 keys=request.form("key")
	 searchid=request.form("searchid")
	 sql="select id,title,times from allarti where "&searchid&" like '%"&keys&"%'"
          else
  if id="" or isnull(id) then
   sql="select id, title,times from allarti order by times desc"
      else
	  sql="select id,title,times from allarti where frominto="&id&" order by times desc"
	   end if
	   end if
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
			  response.write"<script>alert('�벻Ҫ�ڵ�ַ�����������ҳ��');</script>"
			  end if
              rs.absolutepage=currentPage
              page_total=rs.pagecount 
			  do while not rs.eof and maxsize>0 
                maxsize=maxsize-1
    %>
  <tr> 
    <td width="53%"><div align="center"><a href="../listarti.asp?id=<%=rs("id")%>" target="_blank"><%=rs("title")%></a></div></td>
    <td width="27%"><div align="center"><%=rs("times")%></div></td>
    <td width="10%"><div align="center"><a href="editsortnew.asp?id=<%=rs("id")%>">�޸�</a>/<a href="delsortnew.asp?id=<%=rs("id")%>">ɾ��</a></div></td>
    <td width="10%"><div align="center"> 
        <input type="checkbox" name="Board" id="<%=rs("id")%>" style="font-size: 9pt">
      </div></td>
  </tr>
  <%  rs.movenext
      loop
	  else response.write"<tr><td colspan='4'>��������</td></tr>"
	  end if
  %>
  <tr> 
    <td colspan="4"> <div align="center"> &nbsp; <a href="addsortnew.asp">�������</a>&nbsp;&nbsp; 
          <input type="button" value="ȫ ѡ" onclick="sltAll()" name=button1>
        &nbsp;&nbsp; 
        <input type="button" value="�� ��" onclick="sltNull()" name=button2>
        &nbsp;&nbsp; 
        <input type="submit" value="ɾ ��" name="tijiao" onclick="SelectChk('delsortnew.asp')">
      </div></td>
  </tr>
  <tr> 
    <td colspan="4"> 
      <%if page_total>1 then
	       response.write"�ܹ���:"
	       response.Write(page_total)
	       response.write"ҳ."
	       response.write "��ѡ��ҳ��:"
		   if currentPage=1 then
	       response.write"��һҳ"
	       else %>
      <a href="sort_new.asp?page=<%=currentPage-1%>&id=<%=id%>">��һҳ</a> 
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
</script> <select  onChange=GoToWhere(this) name=select5 class="new">
        <%  
	      for i=1 to page_total
	        if i=currentPage then 
		  %>
        <option value="sort_new.asp?page=<%=i%>&id=<%=id%>" selected><%=i%></option>
        <% else %>
        <option value="sort_new.asp?page=<%=i%>&id=<%=id%>"><%=i%></option>
        <%  end if 
			next %>
      </select> 
      <% if currentPage=page_total then
		response.write"��һҳ"
		else %>
      <a href="sort_new.asp?page=<%=currentPage+1%>">��һҳ</a> 
      <% end if
		 else response.write"��������ֻ��һҳ,��ǰҳ��:��1ҳ"
		  end if %>
    </td>
  </tr>
</table>
</form>
<form name="form2" method="post" action="sort_new.asp?action=search" onsubmit="return checkForm(this);">
  <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr>
    <td>
        �ؼ���:<input name="key" type="text" size="15"> &nbsp;&nbsp;
        ���ѷ�ʽ:
        <select name="searchid">
          <option value="title" selected>������</option>
          <option value="body">������</option>
        </select>
        &nbsp;&nbsp;
        <input type="submit" name="Submit" value="����"> </td>
  </tr>
</table>
</form>
<!--#include file="adminfoot.asp"-->
</body>
</html>
