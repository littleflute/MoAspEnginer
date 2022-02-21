<!--#include file="rscoon.asp"-->
<% toptitle="搜索频道" %>
<!--#include file="mztop.asp"-->
<%
   dim keywords,channel,hrefs,listfile,i
        hrefs=""
       keywords=request.Form("keyword")
	   channel=request.Form("channel")
	   select case channel
	   case  "zcfg"
	    sql="select id,title from allarti where title like '%"&keywords&"%'"
		 hrefs="listarti.asp"
		case  "gjzc"
		sql="select id,pytitle from policy where pytitle like'%"&keywords&"%'"
		hrefs="listpy.asp"
		case "mznews"
		sql="select id,nstitle from news where nstitle like '%"&keywords&"%'"
		hrefs="listnews.asp"
		case "spjc"
		sql="select id, vdtitle from video where vdtitle like '%"&keywords&"%'"
		 hrefs="movie.asp"
		case "mztj"
		sql="select tjid,tjtitle from mztj where tjtitle like '%"&keywords&"%'"
         hrefs="dismztj.asp"
		case  "worknet"
		sql="select id,nettitle from worknet where nettitle like '%"&keywords&"%'"
		hrefs="worknet.asp"
		case else
		response.write"<script>alert('对不起,请选择您要搜索的栏目');window.close();</script>"
		 end select
		set rs=conn.execute(sql)
		    if not rs.eof then
		    listfile=rs.getrows
response.write"<TABLE width=766 border=0 align=center cellPadding=0 cellSpacing=0 bgcolor=#FFFFFF><TBODY>"           
				  for i=0 to ubound(listfile,2)
                    response.write"<TR><TD height=20 width='6%'> <DIV align=center><IMG height=10 src='images/rj-tp3.gif' width=6></DIV></TD>"&_
					"<TD class=wz9 height=20 width='94%'><a href='"&hrefs&"?id="&listfile(0,i)&"'>"&listfile(1,i)&"</a></TD></TR>"
				       next
						 else
					   response.write"<tr><td colspan=2>对不起，没有搜索到!!!</td></tr>"
						  end if
					response.write"  </TBODY></TABLE>"
					call mzfoot()
					%>