<!--#include file="conn.asp"-->
<!--#include file="home1.asp"-->
<%tmp = "http://" & request.servervariables("SERVER_NAME") & _            
left(request.servervariables("SCRIPT_NAME"),len(request.servervariables("SCRIPT_NAME"))-len("software.asp"))
   	dim sql
   	dim rs
   	dim classname,classid,Nclassname,Nclassid
	dim lasthits
   	dim title
	if request("id")="" then
		response.write "您没有选择相关软件，请返回"
		response.end
	end if
	set rs=server.createobject("adodb.recordset")
  	sql="select class.class,Nclass.Nclass,download.showname,bb,download.classid,download.Nclassid,download.lasthits from download,class,Nclass where download.classid=class.classid and download.Nclassid=Nclass.Nclassid and download.ID="&request("id")
 	rs.open sql,conn,1,1
 	if not rs.eof then
		showname=rs("showname")
		bb=rs("bb")
		classid=rs("classid")
		Nclassid=rs("Nclassid")
		classname=rs("class")
		Nclassname=rs("Nclass")
		lasthits=rs("lasthits")
 	end if
	rs.close
'更新每周每日数据
	tdate=year(Now()) & "-" & month(Now()) & "-" & day(Now())
	if trim(lasthits)=trim(tdate) then
		sql="update download set dayhits=dayhits+1 where id="&request("id")
		conn.Execute(sql)
'		response.write "success"
	else
		sql="update download set dayhits=1 where id="&request("id")
		conn.Execute(sql)
'		response.write "error"
	end if
	sql="update download set hits=hits+1,lasthits="&tdate&" where ID="&request("id")
	conn.Execute(sql)


	if period_time=<cint(7) then
		sql="update download set weekhits=weekhits+1 where id="&request("id")
		conn.Execute(sql)
	else
		sql="update download set weekhits=1 where id="&request("id")
		conn.Execute(sql)
	end if
%> 
<HTML>
<HEAD>
<META content="text/html; charset=gb2312" http-equiv="Content-Type">
<TITLE><%=showname%></TITLE>
</HEAD>
<!--#include file="head.asp"-->
<table width="770" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#6AAC91"> <div align="center">
        <!--#include file="lanmu.asp"-->
      </div></td>
  </tr>
</table>
<table width="770" border=0 cellspacing="0" cellpadding="0" >
  <tr> 
    <td height="22" bgcolor="#FFFFFF"> 当前位置：<a href="<%=rs1("url")%>"><%=rs1("home")%></a> >> <a href="<%=rs1("urls")%>"><%=rs1("homes")%></a> 
      >> <A href="class.asp?classid=<%=classid%>&Nclassid=<%=Nclassid%>"><%=Nclassname%></A> 
    >> <%=showname%></td>
  </tr>
</table>
<br>
<table   width="770" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr> 
    <td width="170" valign="top"> <table border=0 cellpadding="0" cellspacing="0"  width="168" height="73">
        <tr> 
          <td width="174" height="22" class="li"> 搜索引擎</td>
        </tr>
        <tr bgcolor=#ffffff> 
          <td width="174" align="center" class="3b" background="pic/a-27.gif" bgcolor="#ffffff"><TABLE width="98%" height="63" border=0 align=center cellPadding=0 cellSpacing=2>
              <FORM action=query.asp method=POST name="myfrom">
                <TR> 
                  <TD> <input onfocus="this.value=''" maxLength="50" size="17" title="输入关键字" value="输入关键字" name="keyword" type="text"> 
                  </TD>
                </TR>
                <TR> 
                  <TD><input type="radio" value="title" checked name="action">
                    名称 
                    <input type="radio" name="action" value="content">
                    简介<font color="#FFFFFF">&nbsp;</font></TD>
                </TR>
                <TR> 
                  <TD><input type=submit value="搜 索" name=Submit> <a href="search.asp"> 
                    高级搜索...</a></TD>
                </TR>
              </FORM>
            </TABLE></td>
      </table>
      <br>
      <table width="168"  cellspacing="0" cellpadding="0">
        <tr> 
          <td class="li" height="22"> 本日下载排行</td>
        </tr>
        <tr> 
          <td> 
            <table border="0" cellpadding="4" cellspacing="0">
              <tr> 
                <td valign="top" width="191" class="3b" background="pic/a-27.gif"> 
                  <%                                                                                  
	dim tdate                                                                                  
	tdate=year(Now()) & "-" & month(Now()) & "-" & day(Now())                                                                                  
	sql="select top 10 id,showname,bb from download where "                                                                                  
	sql=sql&" lasthits="&tdate&" and  dayhits>0 "                                                                                  
	sql=sql&" order by dayhits desc"                                                                                  
	rs.open sql,conn,1,1                                                                                  
	if rs.eof and rs.bof then                                                                                  
%>
                  本日没有下载 
                  <%else%>
                  <%do while not rs.eof                                                                                     
response.write "·<A href=software.asp?id="&rs("id")&">"&rs("showname")&"</A><br>"                                                                                     
i=i+1                                                   
if i>=10 then exit do                                                   
rs.movenext                                                 
loop                                                 
i="0"                                                 
	end if                                                 
	rs.close                                                                                         
%>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <br>
      <table width="168" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="li" height="22"> 本周下载排行</td>
        </tr>
        <tr> 
          <td> 
            <table border="0" cellpadding="4" cellspacing="0">
              <tr> 
                <td width="191" class="3b" background="pic/a-27.gif"> 
                  <%                                                                                     
	OldWeek = WeekDay(Date())-1                                                                                     
	If OldWeek = 0 Then OldWeek = 7                                                                                     
	OldWeek = Date()-OldWeek                                                                                     
	NewWeek = Date()+(9-WeekDay(Date()))                                                                                     
	sql="select top 10 id,showname,bb from download where "                                                                                     
	sql=sql&"  (lasthits < " & NewWeek & ") And (lasthits > " & OldWeek & ") and weekhits>0 "                                                                                     
	sql=sql&" order by weekhits desc"                                                                                     
	rs.open sql,conn,1,1                                                                                     
	if rs.eof and rs.bof then                                                                                     
%>
                  本周没有下载 
                  <%else%>
                  <%do while not rs.eof                                                                                     
response.write "·<A href=software.asp?id="&rs("id")&">"&rs("showname")&"</A><br>"                                                                                     
i=i+1                                                
if i>=10 then exit do                                                
rs.movenext                                                
loop                                                
i="0"                                                
	end if                                                 
	rs.close                                                 
%>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td align="right" width="560" bgcolor="#FFFFFF" valign="top"> 
      <%                                          
    sql="select * from download where id="&request("id")                                       
	rs.open sql,conn,1,1                                          
%>
      <table border="0" cellspacing="0" width="100%" cellpadding="3"  class="4b">
        <tr> 
          <td width="14%">软件名称：</td>
          <td colspan="2"><b><%=rs("showname")%></b></td>
        </tr>
        <tr> 
          <td width="14%">软件级别：</td>
          <td width="36%"><b> 
            <% if rs("hy")=1 then
			response.write "<font color=red><b>会员软件</b></font>"
			end if 
                        if rs("hy")=0 then
			response.write "<font color=808080><b>免费软件</b></font>"
			end if %>
            </b></td>
          <td width="50%" rowspan="8"> <script language="javascript" src="../ad/js/softad.js"></script> 
          </td>
        </tr>
        <tr> 
          <td width="14%">软件类型：</td>
          <td width="36%"><%=rs("orders")%>/<%=Nclassname%></td>
        </tr>
        <tr> 
          <td width="14%">运行环境：</td>
          <td width="36%"><%=rs("system")%></td>
        </tr>
        <tr> 
          <td width="14%">软件语言：</td>
          <td width="36%"><%=rs("xxyf")%></td>
        </tr>
        <tr> 
          <td width="14%">软件大小：</td>
          <td width="36%"><%=rs("size")%></td>
        </tr>
        <tr> 
          <td width="14%">软件评价：</td>
          <td width="36%" bgcolor="#FFFFFF"><img src="images/<%=rs("hot")%>star.gif" height="15"></td>
        </tr>
        <tr> 
          <td width="14%">整理日期：</td>
          <td width="36%"><%=rs("dateandtime")%></td>
        </tr>
        <tr> 
          <td width="14%">程序演示：</td>
          <td width="36%"> 
            <%if rs("bb")<>"" then
					  response.write"<a href="&rs("bb")&" target=_blank>点击此处看演示</a>"
					  else
					  response.write"没有演示"
					  end if%>
          </td>
        </tr>
        <tr> 
          <td width="14%">相关链接：</td>
          <td colspan="2"> 
            <%if rs("fromurl")<>"" then
					  response.write"<a href="&rs("fromurl")&" target=_blank>点击此处看演示</a>"
					  else
					  response.write"<a href="&http://www.aaa.com&" target=_blank>点击此处看演示</a>"
					  end if%>
          </td>
        </tr>
        <tr> 
          <td width="14%">下载统计：</td>
          <td colspan="2">本日下载：<%=rs("dayhits")*4%>&nbsp;&nbsp;本周下载：<%=rs("weekhits")*4%>&nbsp;&nbsp;总计下载：<%=rs("hits")*4%></td>
        </tr>
        <%
if rs("filename")="" and rs("filename1")="" and rs("filename2")="" then
%>
        <tr> 
          <td width="14%" >下载地址：</td>
          <td colspan="2">暂时没有下载</td>
        </tr>
        <%
else

if rs("movie")<>"" then
if rs("movie")="win" then     
play="playwin.asp"     
pimg="images/win.gif"     
else     
play="playrm.asp"     
pimg="images/rm.gif"     
end if     
%>
        <script language="JavaScript">     
function windowOpen(loadpos)     
{controlWindow=window.open(loadpos,"surveywin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=610,height=400,status=yes,resizable=yes");     
}     
              </script>
        <tr> 
          <td width="14%" bgcolor="#D8D8D8" >在线播放：</td>
          <td colspan="2" bgcolor="#D8D8D8"><img src="<%=pimg%>">&nbsp; 
            <%if rs("filename")<>"" then%>
            <a href="#" onClick="windowOpen('<%=play%>?id=<%=rs("id")%>&amp;downid=1')">[1]</a> 
            <%end if%>
            <%if rs("filename1")<>"" then%>
            <a href="#" onClick="windowOpen('<%=play%>?id=<%=rs("id")%>&amp;downid=2')">[2]</a> 
            <%end if%>
            <%if rs("filename2")<>"" then%>
            <a href="#" onClick="windowOpen('<%=play%>?id=<%=rs("id")%>&amp;downid=3')">[3]</a> 
            <%end if%>
            <%if rs("filename3")<>"" then%>
            <a href="#" onClick="windowOpen('<%=play%>?id=<%=rs("id")%>&amp;downid=4')">[4]</a> 
            <%end if%>
            <%if rs("filename4")<>"" then%>
            <a href="#" onClick="windowOpen('<%=play%>?id=<%=rs("id")%>&amp;downid=5')">[5]</a> 
            <%end if%>
          </td>
        </tr>
        <%else%>
        <%if rs("filename")<>"" then%>
        <tr bgcolor="#006699"> 
          <td align="center" colspan="3" height="28"> <img border="0" src="images/download2.gif" width="62" height="11">&nbsp;&nbsp;&nbsp; 
            <font color="#FFFFFF">请选择下载地址或类型&nbsp;&nbsp; </font></td>
        </tr>
        <tr bgcolor="#D8D8D8"> 
          <%
if session("user")="" and rs("hy")="1" then
%>
        <tr> 
          <td width="14%" bgcolor="#cdcdcd" > <form method="post" action="users.asp">
              <%
if session("user")<>"" then%>
              <%sql="select * from user where user='"&session("user")&"'"
        rs.open sql,conn,3,3
        %>
              <%rs.close%>
              <%else%>
              会员下载区<img border="0" src="images/jumpback.gif" width="14" height="11"> 
            </form></td>
          <td colspan="2" bgcolor="#CCCCCC"> 
            <input type="password" name="password" size="1" class="input1" style="font-family: Arial; height:0"> 
            <%end if%>
            &nbsp; </td>
        </tr>
        <tr> 
          <%else if rs("filename")<>"" then%>
          <td width="14%" bgcolor="#cdcdcd" height="22"><a href="softdown.asp?downid=1&id=<%=rs("id")%>" target="_blank"><img border="0" src="images/jumpback.gif" width="14" height="11"><%=rs("txtname")%></a></td>
          <td colspan="2" bgcolor="#CCCCCC" height="22"><a href="softdown.asp?downid=1&id=<%=rs("id")%>" target="_blank"><img src="images/download.gif" width=63 height=10 alt="" border=0></a> 
          </td>
        </tr>
        <%end if%>
        <%end if%>
        <%
if session("user")="" and rs("txtname1")="VIP" then
%>
        <%else if rs("filename1")<>"" then%>
        <tr> 
          <td width="14%" bgcolor="#cdcdcd" ><a href="softdown.asp?downid=1&id=<%=rs("id")%>" target="_blank"><img border="0" src="images/jumpback.gif" width="14" height="11"></a><%=rs("txtname1")%></td>
          <td colspan="2" bgcolor="#CCCCCC"><a href="softdown.asp?downid=2&id=<%=rs("id")%>" target="_blank"><img src="images/download.gif" width=63 height=10 alt="" border=0></a> 
          </td>
        </tr>
        <tr> 
          <%end if%>
          <%end if%>
          <%end if%>
          <%if rs("filename2")<>"" then%>
        <tr> 
          <td width="14%" bgcolor="#cdcdcd"><a href="softdown.asp?downid=1&id=<%=rs("id")%>" target="_blank"><img border="0" src="images/jumpback.gif" width="14" height="11"></a><%=rs("txtname2")%></td>
          <td colspan="2" bgcolor="#CCCCCC"><a href="softdown.asp?downid=3&id=<%=rs("id")%>" target="_blank"><img src="images/download.gif" width=63 height=10 alt="" border=0></a> 
          </td>
        </tr>
        <%end if%>
        <%if rs("filename3")<>"" then%>
        <tr> 
          <td width="14%" bgcolor="#cdcdcd"><a href="softdown.asp?downid=1&id=<%=rs("id")%>" target="_blank"><img border="0" src="images/jumpback.gif" width="14" height="11"></a><%=rs("txtname3")%></td>
          <td colspan="2" bgcolor="#CCCCCC"><a href="softdown.asp?downid=4&id=<%=rs("id")%>" target="_blank"><img src="images/download.gif" width=63 height=10 alt="" border=0></a> 
          </td>
        </tr>
        <%end if%>
        <%if rs("filename4")<>"" then%>
        <tr> 
          <td width="14%" bgcolor="#cdcdcd"><a href="softdown.asp?downid=1&id=<%=rs("id")%>" target="_blank"><img border="0" src="images/jumpback.gif" width="14" height="11"></a><%=rs("txtname4")%></td>
          <td colspan="2" bgcolor="#CCCCCC"><a href="softdown.asp?downid=5&id=<%=rs("id")%>" target="_blank"><img src="images/download.gif" width=63 height=10 alt="" border=0></a> 
          </td>
        </tr>
        <%end if%>
        <%end if%>
        <%end if%>
        <tr> 
          <td width="14%" bgcolor="#F5F5F5" height="35"><img border="0" src="images/yes.gif" width="12" height="12">公&nbsp; 
            告：</td>
          <td colspan="2" bgcolor="#F5F5F5" height="35"> 有时网络繁忙或下载人数过多，会出现无法下载的情况，您可以多刷新几次或换个时间再来下载！</td>
        </tr>
        <tr>
          <td bgcolor="#F5F5F5">&nbsp;</td>
          <td bgcolor="#F5F5F5" colspan="2"><script language="javascript" src="../ad/js/banner1.js"></script></td>
        </tr>
        <tr> 
          <td width="14%" bgcolor="#F5F5F5" valign="top"> <img src="images/about.gif" width=9 height=12 alt="" border=0>简　介：</td>
          <td bgcolor="#F5F5F5" colspan="2" valign="top"><%=rs("note")%>　</td>
        </tr>
        <tr> 
          <td bgcolor="#F5F5F5" width="14%" > <p><font color="#FF0000">相关软件：</font> 
          </td>
          <td colspan="2" bgcolor="#F5F5F5"> 
            <%  set rs_xg=server.createobject("adodb.recordset")                                      
    sql_xg="select * from download where id<>"&rs("id")&" and showname like '%"&rs("showname")&"%' order by id desc"                                      
	rs_xg.open sql_xg,conn,1,1                     
	if rs_xg.eof and rs_xg.bof then                                      
response.write ("没有相关软件")                                      
else                                      
do while not rs_xg.eof                                      
response.write "<a href=software.asp?id="&rs_xg("id")&">"&rs_xg("showname")&" "&rs_xg("bb")&"</A><br>"                                      
rs_xg.movenext                                      
loop                                    
end if                                    
rs_xg.close                                     
set rs_xg=nothing%>
          　</td>
        </tr>
        <tr> 
          <td colspan="3"><img border="0" src="images/aboutm.gif" width="15" height="16"> 
            推荐使用：<font color="#FF0000"><b>WinRAR V3.20以上版本</b></font>解压本站软件。<br>
             <img border="0" src="images/aboutm.gif" width="15" height="16"> 
            服务器能力有限，谢绝链接本站软件，如果链接请务必链接软件下载页面，本站不定期跟改下载链接。<br> <img border="0" src="images/aboutm.gif" width="15" height="16"> 
            本站软件都是来自网络，版权属原作者或厂商所有！<br> <img border="0" src="images/aboutm.gif" width="15" height="16"> 
            欢迎广大软件作者以及厂商在本站发布，本站将为您予以能力范围以内的推广！</td>
        </tr>
      </table>
    </td>
  </TR>
</TABLE>
<br>
<!--#include file="foot.asp"-->
</BODY>
<%rs.close                                       
set rs=nothing                                       
conn.close                                                                       
set conn=nothing%>
</HTML><%
'这里定义你的参数就可以了
Session("DownSN")="iwanttohahahahaahahah"
%>