<%
private sub select_page(page_no,total_page)
	'该子程序依次写出各页，并将非当前页设置超链接
	response.write "请选择页码:"
	dim i
	for i=1 to total_page
		if i=page_no then 
			response.write i & "&nbsp"
		else
			response.write "<a href='oneedit.asp?page_no=" & i & "'>" & i & "</a>&nbsp"
		end if
	next
end sub
%>
	