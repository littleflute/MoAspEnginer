<!--#include file="mzconst.asp"-->
<!--#include file="rscoon.asp"-->
<% toptitle="������Ƶ�ļ�" %>
<!--#include file="mztop.asp"-->
		   <script>
 function GoToWhere(s)
 {
 var d = s.options[s.selectedIndex].value;
 window.location=d;
 s.selectedIndex=0;
 }
</script>
  <%  response.write"<table width=766 border=0 align=center cellpadding=0 cellspacing=0>" 
    dim currentpage,id,page_total,i,act
		   set rs = server.createobject("adodb.recordset")
		   sql="select id,vdtitle from video order by times desc"
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
  <tr <% if maxsize mod 2 = 0 then%>bgcolor="#FFFFFF"<%else%>bgcolor="#D6E8E9"<%end if %>> 
    <td width="4%"> 
      <div align="center"><img src="images/rj-tp3.gif" width="10" height="10" align="absmiddle"></div></td>
    <td width="96%" height="20"><a href="movie.asp?id=<%=rs("id")%>"><%=rs("vdtitle")%></a></td>
  </tr>
  <%  rs.movenext
       loop
	   else 
	   response.write"<tr><td>�Բ�����������</td></tr>"
        end if
  response.write"<tr bgcolor=#F7F7F7><td height=20 colspan=2>"
	if page_total>1 then
	       response.write"�ܹ���:["&page_total&"ҳ.��ѡ��ҳ��:"
		   if currentPage=1 then
	       response.write"��һҳ"
	       else
		   response.write"<a href='listallvideo.asp?page="&currentPage-1&"'>��һҳ</a>"
	                   end if  
response.write"<select  onChange=GoToWhere(this) name=select5 class='new'>"
	      for i=1 to page_total
	        if i=currentPage then 
            response.write"<option value='listallvideo.asp?page="&i&"' selected>"&i&"</option>"
            else 
			response.write"<option value='listallvideo.asp?page="&i&"'>"&i&"</option>"
			  end if 
			next 
          response.write"</select>"
		   if currentPage=page_total then
		response.write"��һҳ"
		else 
		response.write"<a href='listallvideo.asp?page="&currentPage+1&"'>��һҳ</a>" 
                         end if
		 else response.write"�����ļ�ֻ��һҳ,��ǰҳ��:��1ҳ"
		  end if 
	response.write"</td></tr><tr bgcolor=#FFFFFF><td height=20 colspan=2>"&_
	"<div align=right><a href='javascript:window.close()'>���رմ��ڡ�</a></div></td></tr></table>" 
      call mzfoot() %>