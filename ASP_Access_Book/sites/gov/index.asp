<!--#include file="mzconst.asp"-->
<!--#include file="rscoon.asp"-->
<!--#include file="totalcount.asp"-->
<% toptitle="政府网欢迎您！" %>
<!--#include file="mztop.asp"-->
<TABLE width=766 border=1 align=center cellPadding=0 cellSpacing=0 bordercolor="#E7E7E7" background="images/tpbg1.gif">
  <TBODY>
  <TR>
      <TD height=33 vAlign=center> 
        <TABLE width="100%" height="100%" border=0 align="center" cellPadding=0 cellSpacing=0>
          <TBODY>
        <TR>
              <TD width="12%" height="31" vAlign=bottom> </TD>
              <TD vAlign=bottom><img height=9 src="images/tpbgtp.gif" 
            width=16> <span><font color="#006633"><%=toppic%></font></span></TD>
              <TD width="20%" vAlign=bottom>
            <marquee direction=left height=5 id=tt6 
                        onMouseOut=tt6.start() onMouseOver=tt6.stop() 
                        scrollamount=1 scrolldelay=60 
                        width=95%>
            </marquee>
            <span class="3dfont">您是第<FONT color=red><%=totalcount%></FONT>位访客</span></TD>
            </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE align=center bgColor=#ffffff border=0 cellPadding=0 cellSpacing=0 
width=766>
  <TBODY>
  <TR>
    <TD background=images/l-bg1.gif height=575 vAlign=top width=187>
      <TABLE align=center border=0 cellPadding=0 cellSpacing=0 height=13 
      width=167>
        <TBODY>
        <TR>
          <TD><IMG height=32 src="images/zcfg1.gif" 
        width=186></TD></TR></TBODY></TABLE>
        <TABLE width=167 border=0 align=center cellPadding=0 cellSpacing=0>
          <TBODY>
        <TR>
              <TD vAlign=top><IMG height=5 src="images/zcfg-tp1.gif" 
            width=100%></TD>
            </TR>
        <TR>
              <TD height=193 vAlign=top> <TABLE 
                  width="99%" border=1 align=center cellPadding=0 cellSpacing=0 bordercolor="#999999">
                  <TBODY>
                    <TR> 
					<% dim allsort  
					sql="select sort,id from allsort order by id"
					      set rs=conn.execute(sql)
						  allsort=rs.getrows
					%>
                      <TD height="188" valign="middle"> &nbsp;&nbsp; <FONT 
                        color=#ff0000>·</FONT><a href="sorts.asp?id=<%=allsort(1,0)%>"><%=allsort(0,0)%></a> <FONT color=#ff0000>·</FONT><a href="sorts.asp?id=<%=allsort(1,1)%>"><%=allsort(0,1)%></a><BR>
                        &nbsp;&nbsp; <FONT color=#ff0000>·</FONT><a href="sorts.asp?id=<%=allsort(1,2)%>"><%=allsort(0,2)%></a> <FONT color=#ff0000>·</FONT><a href="sorts.asp?id=<%=allsort(1,3)%>"><%=allsort(0,3)%></a><BR>
                        &nbsp;&nbsp; <FONT color=#ff0000>·</FONT><a href="sorts.asp?id=<%=allsort(1,4)%>"><%=allsort(0,4)%></a> <FONT color=#ff0000>·</FONT><a href="sorts.asp?id=<%=allsort(1,5)%>"><%=allsort(0,5)%></a><BR>
                        &nbsp;&nbsp; <FONT color=#ff0000>·</FONT><a href="sorts.asp?id=<%=allsort(1,6)%>"><%=allsort(0,6)%></a> <FONT color=#ff0000>·</FONT><a href="sorts.asp?id=<%=allsort(1,7)%>"><%=allsort(0,7)%></a><BR>
                        &nbsp;&nbsp; <FONT color=#ff0000>·</FONT><a href="sorts.asp?id=<%=allsort(1,8)%>"><%=allsort(0,8)%></a> <FONT color=#ff0000>·</FONT><a href="sorts.asp?id=<%=allsort(1,9)%>"><%=allsort(0,9)%></a><BR>
                        &nbsp;&nbsp; <FONT color=#ff0000>·</FONT><a href="sorts.asp?id=<%=allsort(1,10)%>"><%=allsort(0,10)%></a> <FONT color=#ff0000>·</FONT><a href="sorts.asp?id=<%=allsort(1,11)%>"><%=allsort(0,11)%></a><BR>
                        &nbsp;&nbsp; <FONT color=#ff0000>·</FONT><a href="sorts.asp?id=<%=allsort(1,12)%>"><%=allsort(0,12)%></a> <FONT color=#ff0000>·</FONT><a href="sorts.asp?id=<%=allsort(1,13)%>"><%=allsort(0,13)%></a><br>
                        &nbsp;&nbsp; <FONT color=#ff0000>·</FONT><a href="sorts.asp?id=<%=allsort(1,14)%>"><%=allsort(0,14)%></a><BR>
                        &nbsp;&nbsp; <FONT color=#ff0000>·</FONT><a href="sorts.asp?id=<%=allsort(1,15)%>"><%=allsort(0,15)%></a></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
        <TR>
              <TD vAlign=top><IMG height=5 src="images/zcfg-tp3.gif" 
            width=100%></TD>
            </TR></TBODY></TABLE>
      <TABLE border=0 cellPadding=0 cellSpacing=0 height=3 width="100%">
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width="98%">
        <TBODY>
        <TR>
          <TD bgColor=#f7f7f7 height=174 vAlign=top>
            <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
              <TBODY>
              <TR>
                      <td height="30" background="images/zsdwwz.gif"> 
                        <div align="center"><strong><font color="#FFFF00">媒体关注</font></strong></div></td>
                    </TR></TBODY></TABLE>
                <TABLE 
              width="100%" border=1 align=right cellPadding=0 cellSpacing=0 bordercolor="#5698CD">
                  <TBODY>
                    <%  sql="select top 6 nstitle, id from news order by clicks desc"
			       set rs=conn.execute(sql)
				    if not rs.eof then
					   i=6
					do while not rs.eof and i>0 
					  nstitle=rs("nstitle") 
					if len(nstitle)>10 then
					    nstitle=left(nstitle,10)&"..."
						end if
			  %>
                    <TR> 
                      <TD height=20 ><IMG border=0 height=8 
                  src="images/zxzctp4.gif" width=6><a href ="listnews.asp?id=<%=rs("id")%>"><%=nstitle%></a></TD>
                    </TR>
                    <% rs.movenext
				     loop
					 end if
				  %>
                    <TR> 
                      <TD align=left colSpan=2 height=20> <div align="right"><A href="listallns.asp?act=click" target=_blank>更多信息>>..</A> 
                        </div></TD>
                    </TR>
                  </TBODY>
                </TABLE>
		</TD></TR></TBODY></TABLE>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width="98%">
        <TBODY>
        <TR>
          <TD bgColor=#f7f7f7 height=165 vAlign=top>
            <br>
            <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
              <TBODY>
              <TR>
                      <td height="30" background="images/mzdc.gif"> 
                        <div align="center"><strong><font color="#FFFF00">视频点播</font></strong></div></td>
                    </TR></TBODY></TABLE>
                <TABLE 
              width="100%" border=1 align=right cellPadding=0 cellSpacing=0 bordercolor="#78AC55">
                  <TBODY>
                    <% sql="select top 5 id,vdtitle from video order by times desc"
					   set rs=conn.execute(sql)
					   i=5
					   if not rs.eof then
					   do while not rs.eof and i>0
					%>
					<TR> 
                    <TD colSpan=3 height=20><IMG border=0 
                  height=8 src="images/zxzctp4.gif" width=6><a href="movie.asp?id=<%=rs("id")%>" target="_blank"><%=rs("vdtitle")%></a></TD>
                    </TR>
					<%  i=i-1
					   rs.movenext
					   loop
					   else
					   response.write"<tr><td>暂无视频点播</td></tr>"
					   end if
					%>
                    <TR> 
                      <TD width="49%">
<DIV align=right></DIV></TD>
                      <TD width="51%"><div align="right"><a href="listallvideo.asp" target="_blank">更多信息&gt;&gt;..</a></div></TD>
                    </TR>
                  </TBODY>
              </TABLE></TD>
        </TR></TBODY></TABLE>
        <TABLE width="98%" border=1 cellPadding=0 cellSpacing=0 bordercolor="#C98247">
          <TBODY>
        <TR>
          <TD bgColor=#e1f5ff height=90 vAlign=top>
            <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
              <TBODY>
              <TR>
                      <td height="28" background="images/spdb.gif"> 
                        <div align="center"><strong>民政统计</strong></div></td>
                    </TR></TBODY></TABLE>
                <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
                  <TBODY>
				  <%  dim mztj  
				  sql="select top 5 tjtitle,tjid from mztj order by tjtime desc"
				    set rs=conn.execute(sql)
					if not rs.eof then
					  mztj=rs.getrows
				     if ubound(mztj,2)<4 then
				      maxcount=ubound(mztj,2)
				       else
				     maxcount=4
					   end if
				 for i=0 to maxcount
				  %>
                    <TR> 
                      <TD width="19%" height=22> <DIV align=center><IMG height=10 
                  src="images/zczx-tp.gif" width=8></DIV></TD>
                      <TD width="81%" height=22><A href="dismztj.asp?id=<%=mztj(1,i)%>"><%=mztj(0,i)%></A></TD>
                    </TR>
					<% next
					else
					response.write"<tr><td>暂无统计</td></tr>"
					end if
					%>
					<tr>
                      <td colspan="2"><div align="right"><a href="listalltj.asp" target="_blank">更多信息&gt;&gt;..</a></div></td>
                    </tr>
                  </TBODY>
                </TABLE></TD></TR></TBODY></TABLE>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width="98%">
        <TBODY>
        <TR>
          <TD height=130 vAlign=top>
            <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
              <TBODY>
              <TR>
                <TD><IMG height=31 src="images/znjs.gif" 
              width=186></TD></TR></TBODY></TABLE>
            <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
              <TBODY>
              <TR>
                <TD>
                  <FORM action="searchs.asp" method="post" 
                  onsubmit="return checkForm(this);" target="_blank">
                  <TABLE border=0 cellPadding=0 cellSpacing=5 width="100%">
                    <TBODY>
                    <TR>
                      <TD height=25>
                        <DIV align=right>检索范围：</DIV></TD>
                      <TD><SELECT name="channel">
                                    <option value="zcfg">政策法规</option>
                                    <option value="gjzc">国家政策</option>
                                    <option value="mznews">民政新闻</option>
                                    <option value="spjc">视频教程</option>
                                    <option value="mztj">民政统计</option>
                                    <option value="worknet">网上办事</option>
                                    
                                  </SELECT></TD></TR>
                    <TR>
                      <TD>
                        <DIV align=right>关键字：</DIV></TD>
                      <TD><INPUT name="keyword" size=10></TD></TR>
                    <TR>
                      <TD colSpan=2>
                        <DIV align=center><INPUT border=0 height=21 
                        src="images/cx11.gif" type=image width=54> 
                      </DIV></TD></TR></TBODY></TABLE></FORM></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD>
    <TD vAlign=top width=579><TABLE width="100%" border=0 align="center" cellPadding=0 cellSpacing=0>
          <TBODY>
        <TR>
          <TD  width=20></TD>
          <TD vAlign=top width=555>
            <TABLE border=0 cellPadding=0 cellSpacing=0 width=546>
              <TBODY>
              <TR>
                <TD  vAlign=top width=389>
                  <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
                    <TBODY>
                    <TR>
                      <TD background=images/zxzctp3.gif height=23 
                      vAlign=top>
                        <TABLE border=0 cellPadding=0 cellSpacing=0 
width="100%">
                          <TBODY>
                          <TR>
                            <TD 2 width="9%">
                              <DIV align=center><IMG height=12 
                              src="images/zxzctp1.gif" 
width=13></DIV></TD>
                                      <TD width="72%"><strong><font color="#FFFFFF">民政新闻</font></strong></TD>
                                      <TD width="19%"><a href="listallns.asp?act=times"><IMG src="images/more.gif"  width=40 height=13 border="0"></a></TD>
                                    </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
                  <TABLE bgColor=#ebebeb border=0 cellPadding=0 cellSpacing=3 
                  width="100%">
                    <TBODY>
                    <TR>
                      <TD bgColor=#f6f6f6  vAlign=top>
                        <TABLE border=0 cellPadding=0 cellSpacing=0 
width="100%">
                            <TBODY>
                              <% dim news,i,maxcount 
								  sql="select top 12 nstitle,times,id from news order by times desc"
								        set rs=conn.execute(sql)
										news=rs.getrows
								       if ubound(news,2)<8 then
									     maxcount=ubound(news,2)
										 else
										 maxcount=8
										 end if
									   for i=0 to maxcount
								  %>
                              <TR> 
                                <TD  width="10%"> <DIV align=center><IMG height=10 
                              src="images/zczx-tp.gif" width=8></DIV></TD>
                                <TD width="50%"><a href="listnews.asp?id=<%=news(2,i)%>"><%=news(0,i)%></a></TD>
                                <TD  width="30%"><a href="listnews.asp?id=<%=news(2,i)%>">(<%=news(1,i)%>)</a> </TD>
                              </TR>
                              <% next
									%>
                              <TR> 
                                <TD height=20> <DIV align=center></DIV></TD>
                                <TD>&nbsp;</TD>
                                <TD height=20><DIV align=right><A 
                              href="listallns.asp?act=times" target=_blank>更多信息</A> 
                                          <IMG 
                            src="images/xxxx1.gif" width=15 height=15 align="absmiddle"> 
                                          &nbsp;&nbsp;</DIV></TD>
                              </TR>
                            </TBODY>
                          </TABLE></TD></TR></TBODY></TABLE>
                  <br>
                  <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
                    <TBODY>
                    <TR>
                      <TD background=images/zxzctp3.gif height=23 
                      vAlign=top>
                        <TABLE border=0 cellPadding=0 cellSpacing=0 
width="100%">
                          <TBODY>
                          <TR>
                            <TD height=22 width="9%">
                              <DIV align=center><IMG height=12 
                              src="images/zxzctp1.gif" 
width=13></DIV></TD>
                                      <TD width="72%"><strong><font color="#FFFFFF">热点透视</font></strong></TD>
                                      <TD width="19%"><a href="sorts.asp"><IMG src="images/more.gif" width=40 height=13 border="0"></a></TD>
                                    </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
                  <TABLE bgColor=#ebebeb border=0 cellPadding=0 cellSpacing=3 
                  width="100%">
                    <TBODY>
                    <TR>
                              <TD bgColor=#f6f6f6  vAlign=top> 
                                <TABLE border=0 cellPadding=0 cellSpacing=0 
width="100%">
                            <TBODY>
                              <%  dim hots  
								  sql="select top 6 id,title,clickcount from allarti order by clickcount desc"
								          set rs=conn.execute(sql)
										  	hots=rs.getrows
								         if ubound(hots,2)<5 then
									     maxcount=ubound(hots,2)
										 else
										 maxcount=5
										 end if
									   for i=0 to maxcount
								  %>
                              <TR> 
                                <TD height=20 width="10%"> <DIV align=center><IMG height=10 
                              src="images/zczx-tp.gif" width=8></DIV></TD>
                                <TD height=20 width="71%"><a href="listarti.asp?id=<%=hots(0,i)%>"><%=hots(1,i)%></a> </TD>
                                <TD width="19%">(<font color="#FF0000">人气:<%=hots(2,i)%></font>)</TD>
                              </TR>
                              <% next %>
                              <TR> 
                                <TD height=20> <DIV align=center></DIV></TD>
                                <TD height=20 colspan="2"><DIV align=right><a href="sorts.asp">更多信息</a> 
                                          <IMG src="images/xxxx1.gif" width=15 height=15 align="absmiddle"> 
                                          &nbsp;&nbsp;</DIV></TD>
                              </TR>
                            </TBODY>
                          </TABLE></TD></TR></TBODY></TABLE><BR>
                  <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
                          <TBODY>
                    <TR>
                              <TD background=images/zxzctp3.gif height=25> 
                                <TABLE border=0 cellPadding=0 cellSpacing=0 
width="100%">
                          <TBODY>
                          <TR>
                            <TD width="9%">
                              <DIV align=center><IMG height=12 
                              src="images/zxzctp1.gif" 
width=13></DIV></TD>
                                      <TD width="72%"><strong><font color="#FFFFFF">民政相关链接</font></strong></TD>
                                      <TD 
                    width="19%"><img src="images/more.gif" width="40" height="13" border="0"></TD>
                                    </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
                  <TABLE bgColor=#ebebeb border=0 cellPadding=0 cellSpacing=3 
                  width="100%">
                    <TBODY>
                    <TR>
                      <TD bgColor=#f6f6f6  vAlign=top>
                        <TABLE border=0 cellPadding=0 cellSpacing=0 
width="100%">
                          <TBODY>
                          <TR>
                            <TD vAlign=top width="48%">
                              <TABLE border=0 cellPadding=0 cellSpacing=0 
                              width="100%">
                                <TBODY>
                                  <% 
										     dim links,j 
											   j=0
										     sql="select webaddr,zwname from aboutlink"
											 set rs=conn.execute(sql)
											 links=rs.getrows
											 for i=0 to ubound(links,2)
										      if j=2 then
											  response.write"<tr>"
										          j=1
												  else
												  j=j+1
												  end if
										  %>
                                          <TD height=20 width="30"> 
                                              <DIV align=center><IMG height=11 
                                src="images/tp1.gif" width=11></DIV></TD>
                                <TD height=20><A 
                                href="<%=links(0,i)%>" target=_blank><%=links(1,i)%></A> </TD>
                                <% next %>
                              </TABLE></TD>
                          </TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD>
                <TD vAlign=top width=173>
                  <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
                    <TBODY>
                    <TR>
                      <TD width="12%">&nbsp;</TD>
                              <TD width="88%"><IMG src="images/wsbs.gif" 
                        width=100%></TD>
                            </TR>
                    <TR>
                      <TD >&nbsp;</TD>
                      <TD vAlign=top>
                        <TABLE 
                        width="97%" border=0 align=center cellPadding=0 cellSpacing=0 bgcolor="#ededed">
                            <TBODY>
							   <% dim works  
							   sql="select top 8 id,nettitle from worknet order by nettime desc"
							     set rs=conn.execute(sql)
								    if not rs.eof then
								   works=rs.getrows
							        if ubound(works,2)<8 then
									     maxcount=ubound(works,2)
										 else
										 maxcount=8
										 end if
									   for i=0 to maxcount
							   %>
                              <TR> 
                                <TD height=20><SPAN ><a href="worknet.asp?id=<%=works(0,i)%>"> 
                                  <li><%=works(1,i)%></a></SPAN></TD>
                              </TR>
							  <% next
							  else
							  response.write"<tr><td><li>暂无内容</td></tr>" 
							    end if%>
                              <TR> 
                                <TD height=10>
<div align="right">
  <p><a href="listworknet.asp" target="_blank">--更多<br>
    <br>
    
    </a></p>
  </div></TD>
                              </TR>
                            </TBODY>
                          </TABLE></TD></TR></TBODY></TABLE>
                  <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
                    <TBODY>
                    <TR>
                      <TD width="12%">&nbsp;</TD>
                              <TD width="88%"><IMG height=32 
                        src="images/mjzzgg.gif" width=100%></TD>
                            </TR>
                    <TR>
                      <TD >&nbsp;</TD>
                      <TD vAlign=top>
                        <TABLE 
                        width="97%" border=0 align=center cellPadding=0 cellSpacing=0 bgcolor="#ededed">
                                  <TBODY>
							<%  dim   mjzj
							sql="select top 7 title,id from allarti where frominto=15 order by times desc"
							      set rs=conn.execute(sql)
								  if not rs.eof then
								   mjzj=rs.getrows 
							     if ubound(mjzj,2)<7 then
									     maxcount=ubound(mjzj,2)
										 else
										 maxcount=4
										 end if
									   for i=0 to maxcount 
							%>
                              <TR> 
                                <TD height=20>&nbsp;<a href="listarti.asp?id=<%=mjzj(1,i)%>"><%=mjzj(0,i)%></a></TD>
                              </TR>
							  <% next
							  else
							  response.write"<tr><td>暂无记录</td></td>"
							    end if
							  %>
                              <TR> 
                                <TD height=10> 
                                  <DIV 
                    align=right><a href="sorts.asp?id=15" target="_blank">--更多</a></DIV></TD>
                              </TR>
                            </TBODY>
                          </TABLE></TD></TR></TBODY></TABLE>
                  <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
                    <TBODY>
                    <TR>
                      <TD width="12%">&nbsp;</TD>
                              <TD width="88%"><IMG height=32 
                        src="images/wsdc.gif" width=100%></TD>
                            </TR>
                    <TR>
                      <TD >&nbsp;</TD>
                      <TD  vAlign=top>
                        <TABLE bgColor=#f6f6f6 border=0 cellPadding=0 
                        cellSpacing=0 width="100%">
                                  <FORM action="votes.asp?action=poll" method=post>
                                    <INPUT name=voteID type=hidden value=6>
                                    <TBODY>
                                      <TR> 
                                        <%  dim stat,statid
						     sql="select title,id from stat where isvote='true'"
							 set rs=conn.execute(sql)
							     if not rs.eof then
								 stat=rs("title")
								 statid=rs("id")
								 else
								 stat="暂无调查信息"
								 statid=""
								 end if
						  %>
                                        <TD align="center"> &nbsp;&nbsp;<%=stat%> <input type="hidden" name="statid" value="<%=statid%>">
                                        </TD>
                                      </TR>
                                      <TR> 
                                        <TD> <TABLE border=0 cellPadding=0 cellSpacing=0 
                              width="100%">
                                            <TBODY>
                                              <%  if statid<>"" then  
										sql="select voteid,votetitle from votes where fromid="&statid&" order by counts desc"
							                   set rs=conn.execute(sql)
											   if not rs.eof then
											   do while not rs.eof 
										 %>
                                              <TR> 
                                                <TD height=20 width="29%"> <DIV align=center> 
                                                    <INPUT name=rv type=radio value=<%=rs("voteid")%>>
                                                  </DIV></TD>
                                                <TD height=20 width="71%"><%=rs("votetitle")%></TD>
                                              </TR>
                                              <%    rs.movenext
										      loop
											  else
											  response.write"<tr><td colspan='2'>暂无调查需要</td></tr>"
										      end if
										%>
                                            </TBODY>
                                          </TABLE></TD>
                                      </TR>
                                      <TR valign="top"> 
                                        <TD> <TABLE border=0 cellPadding=0 cellSpacing=0 
                              width="100%">
                                            <TBODY>
                                              <TR> 
                                                <TD vAlign=top></TD>
                                                <TD></TD>
                                              </TR>
                                              <TR> 
                                                <TD vAlign=top> <DIV align=center> 
                                                    <INPUT height=21  src="images/tp.gif" type=image width=60>
                                                  </DIV></TD>
                                                <TD vAlign=top> <DIV align=center><a href="#" onClick="window.open('votes.asp?action=look','','width=400,height=300,top='+(screen.availHeight-240)/2+',left='+(screen.availWidth-400)/2)"><IMG border=0 height=21 src="images/ck.gif" width=60></a> 
                                                  </DIV></TD>
                                              </TR>
                                              <% end if %>
                                            </TBODY>
                                          </TABLE></TD>
                                      </TR>
                                    </TBODY>
                                  </FORM>
                                </TABLE></TD></TR></TBODY></TABLE>
                  </TD>
              </TR></TBODY></TABLE>
            </TD>
        </TR></TBODY></TABLE>
      <br>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
        <TBODY>
          <tr>
            <td height="27" colspan="2"   background="images/rj-bg4.gif"><div align="center"><strong>最新国家政策</strong></div></td>
          </tr>
          <TR>
            <%  dim pytitle,pybody,nstitle
							     sql="select top 1 id,pytitle,pybody,pyimg from policy order by times desc"
								  set rs=conn.execute(sql)
								  if not rs.eof then
								  pytitle=rs("pytitle")
								  pybody=rs("pybody")
								  pyid=rs("id")
								  if rs("pyimg")="" or isnull(rs("pyimg")) then
								    pyimg="images/noname.gif"
								    else
								  pyimg="admingly/movie/"&rs("pyimg")
								    end if
								 if len(pybody)>200 then
								    pybody=left(pybody,200)&"...."
									end if
								  else
								  pytitle="暂无规定"
								  pybody="暂无内容"
								  pyimg=""
								  end if
								  rs.close
							  %>
            <TD>&nbsp;<IMG 
                        src="<%=pyimg%>" alt="q" width=150 height=120 border="0"></TD>
            <TD align=left height=145 width="100%"><br>
                <div align="center"><%=pytitle%></div>
              <br>
              &nbsp;&nbsp;<a href="listpy.asp?id=<%=pyid%>" target="_blank"><%=pybody%></a> <BR>
              <DIV align=right><A href="listallpy.asp" target=_blank>更多信息</A> <IMG src="images/xxxx1.gif" alt="w3" width=15 height=15 align="absmiddle"> &nbsp;&nbsp;</DIV></TD>
          </TR>
        </TBODY>
      </TABLE>
      <TABLE 
                              width="100%" height="100" border=0 align="left" cellPadding=0 cellSpacing=0 bordercolor="#FFFFFF">
        <TBODY>
          <tr>
            <td width="24%" height="22" background="images/title_7_1.gif">　　　<strong>站内公告</strong></td>
            <td width="76%">&nbsp;</td>
          </tr>
          <TR>
            <TD height="72" colspan="2" vAlign=bottom><table width="100%" height="100%" border="1" cellpadding="1" cellspacing="1" bordercolor="#CCCCCC">
                <tr>
                  <td><% 
											  sql="select top 1 title,body from affiche order by aftime desc"
											      set rs=conn.execute(sql)
												  if not rs.eof then
											  %>
                      <marquee direction=up height=100% id=tt666 
                        onMouseOut=tt666.start() onMouseOver=tt666.stop() 
                        scrollamount=1 scrolldelay=10 
                        width=95%>
                      <div align="center"><%=rs("title")%></div>
                        <br>
                      <%=rs("body")%>
                      </marquee>
                      <% end if %>
                  </td>
                </tr>
            </table></TD>
          </TR>
        </TBODY>
      </TABLE></TD></TR></TBODY></TABLE>
                       <% call mzfoot() %>
