<!--#include file="rscoon.asp"-->
<% toptitle="请选择您要观看的内容" %>
<!--#include file="mztop.asp"-->
<!--#include file="mzconst.asp"-->
<TABLE align=center bgColor=#ffffff border=0 cellPadding=0 cellSpacing=0 
width=766>
  <TBODY>
  <TR>
    <TD background="" height=453 vAlign=top width=186>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
        <TBODY>
        <TR>
          <TD height=38 vAlign=top>
            <TABLE border=0 cellPadding=0 cellSpacing=0 width=180>
              <TBODY>
              <TR>
                      <TD background=images/tpbg.gif 
              height=38>&nbsp;</TD>
                    </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
        <TBODY>
        <TR>
              <TD><IMG height=44 src="images/rj-zcfgtp.gif" 
        width=180></TD>
            </TR></TBODY></TABLE>
        <TABLE width=180 height="25" border=0 cellPadding=0 cellSpacing=0>
          <TR> 
            <TD bgColor=#f9f9f9>
              <!--#include file="leftmz.asp"--></TD>
          </TR>
        </TABLE></TD>
    <TD bgColor=#f9f9f9 vAlign=top width=580>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
        <TBODY>
        <TR>
          <TD background="">&nbsp;</TD></TR></TBODY></TABLE>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
        <TBODY>
        <TR>
          <TD>
            <DIV align=right><IMG height=60 src="images/rj-zcfg.gif" 
            width=580></DIV></TD></TR></TBODY></TABLE>
        <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
          <TBODY>
        <TR>
              <TD background=images/rj-bg4.gif bgColor=#ffffff height=23> 
                <TABLE border=0 cellPadding=0 cellSpacing=0 width="98%">
                  <TBODY>
              <TR>
                <TD class=wz6 height=14 width="4%"></TD>
                      <TD class=wz6 vAlign=bottom width="96%">当前位置：<a href="index.asp">首页</a> 
                        &gt;&gt; <FONT 
                  color=#663300><a href="sorts.asp">政策法规</a></FONT></TD>
                    </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
        <TBODY>
        <TR>
          <TD background="" height=5></TD></TR></TBODY></TABLE>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
        <TBODY>
        <TR>
          <TD>&nbsp;</TD></TR></TBODY></TABLE>
        <TABLE width="100%" border=0 align="center" cellPadding=0 cellSpacing=0>
          <TBODY>
        <TR>
              <TD height=334 width="6%">&nbsp;</TD>
          <TD vAlign=top width="94%">
            <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
                  <TBODY>
		    <%   dim currentpage,id,page_total,i
			 id=request.querystring("id")
       	set rs = server.createobject("adodb.recordset")
		    if id="" or isnull(id) then
			sql="select id,title,frominto,times from allarti order by clickcount desc"
                 else
				 if instr(id,"or")>0 or instr(id,"and")>0 or isnumeric(id)<>true then
				 response.redirect"admingly/error.htm"
                      else 
		   sql="select id,title,frominto,times from allarti where frominto="&id&" order by times desc"
		         end if
				 end if	
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
			  end if
              rs.absolutepage=currentPage
              page_total=rs.pagecount 
			  do while not rs.eof and maxsize>0 
                maxsize=maxsize-1
				  %>
                    <TR><TD height=20 width="6%"> <DIV align=center><IMG height=10 src="images/rj-tp3.gif" width=6></DIV></TD>
                      <TD class=wz9 height=20 width="94%"><a href="listarti.asp?id=<%=rs("id")%>" target="_blank"><%=rs("title")%></a><% if day(rs("times"))=day(now) then%>&nbsp;<IMG height=10 
                  src="images/hot.gif" width=20><% end if %></TD>
                    </TR>
					<%  rs.movenext
					     loop 
						 else
					   response.write"<tr><td>对不起，此版块暂无内容</td></tr>"
					      end if
					response.write"<tr><td colspan=2>"
					   if page_total>1 then
	       response.write"总共有:["&page_total&"]页.请选择页码:"
	       if currentPage=1 then
	       response.write"上一页"
	       else 
		   response.write"<a href='sorts.asp?page="&currentPage-1&"&id="&id&"'>上一页</a>" 
	                   end if  
	              for i=1 to page_total
		          if i=currentPage then 
			      response.write i & "&nbsp"
		                   else
			   response.write "<a href='sorts.asp?page=" & i & "&id="&id&"'>" & i & "</a>&nbsp"
		end if
	next
	 if currentPage=page_total then
		response.write"下一页"
		else response.write"<a href='sorts.asp?page="&currentPage+1&"&id="&id&"'>下一页</a>" 
        end if
	end if 
   response.write"</td></tr></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>"
	 call mzfoot() %>