<%
private sub select_page(page_no,total_page)
	'���ӳ�������д����ҳ�������ǵ�ǰҳ���ó�����
	response.write "��ѡ��ҳ��:"
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
	