<!--#include file="rscoon.asp"-->
<!--#include file="totalcount.asp"-->
<% toptitle="查看咨询" %>
<!--#include file="mztop.asp"-->
  <% 
  response.write"<table width=766  border=0 align=center cellpadding=0 cellspacing=0 bordercolordark=#FFFFFF bgcolor=#EAEAEA>"
    dim currentpage,maxsize,page_total,i,id,contacts,j,allcount,maxcontact
	        maxsize=8
			id=request.QueryString("id")
			if id="" or isnull(id) then 
            sql="select * from contact order by times desc"
			    else
				if isnumeric(id)<>true or instr(id,"or")>0 or instr(id,"and") then
				response.Redirect"admingly/error.htm"
				else
			sql="select * from contact where id="&id
				end if
				end if
			set rs = server.createobject("adodb.recordset")
			rs.open sql,conn,1,1
			if not rs.eof then
			 if not isempty(request("page")) then
      		currentPage=cint(request("page"))
               	else
      		currentPage=1
             	end if
              contacts=rs.getrows
                      rs.close  
					  else
					  response.Redirect"admingly/error.htm"
					  end if
			sql="select count(*) as al from contact"
			      set rs=conn.execute(sql)
			          allcount=rs("al")
					  if allcount mod maxsize =0 then
					  page_total=allcount/maxsize
					     else
					page_total=(allcount\maxsize)+1
						 end if
						 if  currentpage>page_total then
						 currentpage=page_total
						 end if
				if allcount>maxsize*currentpage then
				 maxcontact=maxsize*currentpage-1
				          else
						  maxcontact=allcount-1
						  end if		             
			  for j=maxsize*(currentpage-1) to maxcontact
 response.write"<tr><td height=30 colspan=3 background='images/dhbg2.gif'><div align='center'><strong>来自"&contacts(2,j)&"的朋友的咨询</strong></div></td></tr>"&_
"<tr><td width=30% height=50 rowspan=2  align=center valign=top> <table width=100% border=0><tr><td width=35% height=23>留言者:</td>"&_
"<td width='65%'>"&contacts(1,j)&"</td></tr><tr><td height='27'>联系电话:</td><td>"&contacts(4,j)&"</td></tr>"&_
"<tr><td height=24>手机号码:</td><td>"&contacts(5,j)&"</td></tr><tr><td height='24'>家庭住址:</td><td>"&contacts(3,j)&"</td></tr></table></td>"&_
"<td width='0%' rowspan='2'  align='center' valign='top' background='images/dhtp2.gif'>&nbsp;</td><td width='70%' height='74' valign=top>咨询内容：<br> &nbsp;&nbsp;&nbsp;"&contacts(6,j)&"</td></tr>"&_
"<tr class='p9'><td height='58' valign=top><font color='#FF0000'>咨询回复:</font>"
if contacts(8,j)="" or isnull(contacts(8,j)) then response.write"正在处理中...请明天再来查看" else response.write(contacts(8,j)) end if
response.write"</td></tr>"                   
            next
  
  response.write"<tr><td height=20 colspan=4 valign='top'>" 
       if page_total>1 then
	response.write"总共有:"&(page_total)&"页.请选择页码:"
	if currentPage=1 then
	response.write"上一页"
	else 
	 response.write"<a href='discontact.asp?page="&currentPage-1&"'>上一页</a>" 
	end if  
	for i=1 to page_total
		if i=currentPage then 
			response.write i & "&nbsp"
		else
			response.write "<a href='discontact.asp?page=" & i & "'>" & i & "</a>&nbsp"
		end if
	next
	 if currentPage=page_total then
		response.write"下一页"
		else  response.write"<a href='discontact.asp?page="&currentPage+1&"'>下一页</a>" 
        end if
	 end if 
	 response.write"</td></tr></table>"
	 call mzfoot()
	 %>

