<!--#include file="admintop.asp"-->
<!--#include file="../mzconst.asp"-->
<!--#include file="isadmin.asp"-->
<form name="form1" method="POST">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr bgcolor="#B1E1F3"> 
      <td height="30" colspan="4"> 
        <div align="center"><strong>������Ƶ�ļ�</strong></div></td>
  </tr>
  <tr> 
    <td width="53%"><div align="center">��Ƶ�ļ�</div></td>
    <td width="27%"><div align="center">�ϴ�����</div></td>
    <td width="10%"><div align="center">����</div></td>
    <td width="10%"><div align="center">ѡ��</div></td>
  </tr>
  <%   dim currentpage,id,page_total,i 
   sql="select id, vdtitle,times from video order by times desc"
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
    <td width="53%"><div align="center"><%=rs("vdtitle")%></div></td>
    <td width="27%"><div align="center"><%=rs("times")%></div></td>
    <td width="10%"><div align="center"><a href="editvideo.asp?id=<%=rs("id")%>">�޸�</a>/<a href="delvd.asp?id=<%=rs("id")%>">ɾ��</a></div></td>
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
    <td colspan="4"> <div align="center">
          <input type="button" name="addnew" value="�����֪����Ƶ�ļ�" onclick="BoardWin('addnew.asp')">
         &nbsp;&nbsp;<input type="button" value="�ϴ���Ƶ�ļ�" onclick="BoardWin('upload.asp')" name=add>
        &nbsp;&nbsp;
        <input type="button" value="ȫ ѡ" onclick="sltAll()" name=button1>
        &nbsp;&nbsp;
        <input type="button" value="�� ��" onclick="sltNull()" name=button2>
        &nbsp;&nbsp;
        <input type="submit" value="ɾ ��" name="tijiao" onclick="SelectChk('delvd.asp')">
      </div></td></tr>
  <tr> 
    <td colspan="4">
	<%if page_total>1 then
	       response.write"�ܹ���:"
	       response.Write(page_total)
	       response.write"ҳ."
	       response.write "��ѡ��ҳ��:"
		   if currentPage=1 then
	       response.write"��һҳ"
	       else %> <a href="video.asp?page=<%=currentPage-1%>">��һҳ</a> 
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
            <option value="video.asp?page=<%=i%>" selected><%=i%></option>
            <% else %>
			<option value="video.asp?page=<%=i%>"><%=i%></option>
			<%  end if 
			next %>
          </select>
		  <% if currentPage=page_total then
		response.write"��һҳ"
		else %> <a href="video.asp?page=<%=currentPage+1%>">��һҳ</a> 
                        <% end if
		 else response.write"��������ֻ��һҳ,��ǰҳ��:��1ҳ"
		  end if %></td>
  </tr>
</table>
</form>
</body>
</html>
