<%
private sub select_page(page_no,total_page)
	'该子程序依次写出各页，并将非当前页设置超链接
	response.write "请选择页码:"
	dim i
	for i=1 to total_page
		if i=page_no then 
			response.write i & "&nbsp"
		else
			response.write "<a href='board.asp?boardid="&boardid&"&page_no=" & i & "'>" & i & "</a>&nbsp"
		end if
	next
end sub

function gotTopic(str,strlen)
	dim l,t,c, i
	l=len(str)
	t=0
	for i=1 to l
	c=Abs(Asc(Mid(str,i,1)))
	if c>255 then
	t=t+2
	else
	t=t+1
	end if
	if t>=strlen then
	gotTopic=left(str,i)&"..."
	exit for
	else
	gotTopic=str
	end if
	next
end function
banq ="Powered by：<a href=http://www.whuhp.org> whuhp.org </a><br>"
%>
	