<!--#include file="rscoon.asp"-->
<% toptitle="��������" %>
<!--#include file="mztop.asp"-->
<!--#include file="mzconst.asp"-->
		   <script>
 function GoToWhere(s)
 {
 var d = s.options[s.selectedIndex].value;
 window.location=d;
 s.selectedIndex=0;
 }
</script>
  <% response.write"<table width=766 border=0 align=center cellpadding=0 cellspacing=0>" 
    dim currentpage,id,page_total,i,colors,listns,allcount,maxns,j,rs1
			 if not isempty(request.querystring("page")) then
      		currentPage=cint(request.querystring("page"))
         	else
      		currentPage=1
        	end if  
		   sql="select id,nettitle from worknet order by nettime desc"
				set rs=conn.execute(sql)
				if not rs.eof then
				listns=rs.getrows
				sql="select count(*) as al from worknet "
				 set rs=conn.execute(sql)
				   allcount=rs("al")
				  else
			 response.write"<tr><td>�Բ�����������</td></tr>"
                   end if
			       if allcount mod maxsize = 0 then
				    page_total=allcount/maxsize
				   else
				    page_total=cint(allcount\maxsize) +1
				   end if
			  if currentpage>page_total then
			  currentpage=allpage
			  response.write"<script>alert('�벻Ҫ�ڵ�ַ�����������ҳ��');</script>"
			  end if
			 if maxsize*currentpage>allcount then
			 maxns=allcount
                else
				maxns=maxsize*currentpage
				  end if
				 for j=maxsize*(currentpage-1) to maxns-1
				 if j mod 2=0 then colors="#FFFFFF" else colors="#D6E8E9" end if
  response.write"<tr bgcolor="&colors&">"&_
  "<td width='4%'><div align=center><img src=images/rj-tp3.gif width=10 height=10 align=absmiddle></div></td>"&_
  "<td width=96% height=20><a href='worknet.asp?id="&listns(0,j)&"'>"&listns(1,j)&"</a></td></tr>"
        next
response.write"<tr bgcolor=#F7F7F7><td height=20 colspan=2>"
      if page_total>1 then
	       response.write"�ܹ���:["&page_total&"]ҳ.��ѡ��ҳ��:"
		   if currentPage=1 then
	       response.write"��һҳ"
	       else 
		   response.write"<a href='listworknet.asp?act="&act&"&page="&currentPage-1&"'>��һҳ</a>" 
	                   end if
   response.write"<select  onChange=GoToWhere(this) name=select5>"
	      for i=1 to page_total
	        if i=currentPage then 
            response.write"<option value='listworknet.asp?act="&act&"&page="&i&"' selected>"&i&"</option>"
            else 
			response.write"<option value='listworknet.asp?act="&act&"&page="&i&"'>"&i&"</option>"
			 end if 
			next 
          response.write"</select>"
		  if currentPage=page_total then
		  response.write"��һҳ"
		 else 
		response.write"<a href='listworknet.asp?page="&currentPage+1&"'>��һҳ</a>" 
        end if
		 else response.write"��������ֻ��һҳ,��ǰҳ��:��1ҳ"
		  end if 
	response.write"</td></tr><tr bgcolor=#FFFFFF><td height=20 colspan=2><div align=right><a href='javascript:window.close()'>���رմ��ڡ�</a></div></td>"&_
	"</tr></table>"
     call mzfoot()
%>