<!--#include file="conn.asp"-->
<!--#include file="home1.asp"-->
<%
  dim rs
  set rs = server.createobject("adodb.recordset")
%>
<html>
<HEAD>
<title><%=rs1("home")%>---<%=rs1("homes")%></title>
<META http-equiv="Content-Language" content="zh-cn">
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="head.asp"-->
<table width="770" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#6AAC91"> <div align="center">
        <!--#include file="lanmu.asp"-->
      </div></td>
  </tr>
</table>
<TABLE width="770" border="0" cellPadding="0" cellSpacing="0" bgcolor="#FFFFFF">
  <TR align="middle"> 
    <TD vAlign="top" width="175"> 
      <table border=0 cellpadding="0" cellspacing="0"  width="168" height="73">
        <tr>
          <td width="174" height="22" class="li"> 搜索引擎</td>
        </tr>
        <tr bgcolor=#ffffff>
          <td width="174" align="center" class="3b" background="pic/a-27.gif" bgcolor="#ffffff"><TABLE width="98%" height="63" border=0 align=center cellPadding=0 cellSpacing=2>
              <FORM action=query.asp method=POST name="myfrom">
                <TR>
                  <TD><input onfocus="this.value=''" maxLength="50" size="17" title="输入关键字" value="输入关键字" name="keyword" type="text">
                  </TD>
                </TR>
                <TR>
                  <TD><input type="radio" value="title" checked name="action">
              名称
                <input type="radio" name="action" value="content">
              简介<font color="#FFFFFF">&nbsp;</font></TD>
                </TR>
                <TR>
                  <TD><input type=submit value="搜 索" name=Submit>
                      <a href="search.asp"> 高级搜索...</a></TD>
                </TR>
              </FORM>
          </TABLE></td>
      </table>
      <br>
      <table border=0 cellpadding="0" cellspacing="0"  width="168">
        <tr> 
          <td class="li" height="22"> 本日下载排名</td>
        </tr>
        <tr> 
          <td width="181" height="34" class="3b" background="pic/a-27.gif"> 
            <%
	dim tdate
	tdate=year(Now()) & "-" & month(Now()) & "-" & day(Now())
	sql="select top 6 id,showname,bb,dayhits from download where "
	sql=sql&" lasthits="&tdate&" and dayhits>0 "
	sql=sql&" order by dayhits desc"

		rs.open sql,conn,1,1
	if rs.eof and rs.bof then
response.write ("本日没有下载")
else
do while not rs.eof
response.write ("·<A href=software.asp?id="&rs("id")&" title=今日下载："&rs("dayhits")&"次>"&rs("showname")&"</A><br>")
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
      <br>
      <table border=0 cellpadding="0" cellspacing="0"  width="168">
        <tr> 
          <td class="li" height="22"> 本周下载排名</td>
        </tr>
        <tr> 
          <td width="181" background="pic/a-27.gif" class="3b"> 
            <%                                                                                     
	OldWeek = WeekDay(Date())-1                                                                                     
	If OldWeek = 0 Then OldWeek = 7                                                                                     
	OldWeek = Date()-OldWeek                                                                                     
	NewWeek = Date()+(9-WeekDay(Date()))                                                                                     
	sql="select top 8 id,showname,bb from download where "                                                                                     
	sql=sql&"  (lasthits < " & NewWeek & ") And (lasthits > " & OldWeek & ") and weekhits>0 "                                                                                     
	sql=sql&" order by weekhits desc"                                                                                   
	
	rs.open sql,conn,1,1                                                                                     
	if rs.eof and rs.bof then                                                                                     
%>
            <p>本周没有下载 </p>
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
        </tr>
      </table>
      <br>
      <table border=0 cellpadding="0" cellspacing="0"  width="168">
        <tr> 
          <td width="181" class="li" height="22"> 全部下载排名</td>
        </tr>
        <tr> 
          <td width="181" background="pic/a-27.gif" class="3b"> 
            <%                   
	sql="select top 10 id,showname,bb,hits,dayhits from download "                   
	sql=sql&" order by hits desc"                   
	rs.open sql,conn,1,1                   
	if rs.eof and rs.bof then
response.write ("没有下载")
else
do while not rs.eof
response.write ("·<A href=software.asp?id="&rs("id")&" title=总下载："&rs("hits")*3&"次&#13;&#10;本日下载："&rs("dayhits")&"次>"&rs("showname")&"</A><br>")
	rs.movenext                                     
	loop                                     
	end if                                     
	rs.close                                     
%>
        </tr>
      </table>
    </TD>
    <TD vAlign="top" align="center"> 
      <table width="95%">
        <tr> 
          <td> 
            <%
dim sql1,totalsoft,foolcat
sql1="select count(id) from download"
totalsoft=conn.execute(sql1)(0)
foolcat = "本栏目有源码：<font color=#ff0000><b>" & totalsoft*4&"</b></font>个"
response.write ""&foolcat
%>
            <a href="rjfl.asp">软件分类</a>&nbsp; <a href="ph.asp">下载排行</a>&nbsp; 
            <a href="zx.asp">最新更新</a>&nbsp; <a href="tj.asp">推荐排行</a></td>
        </tr>
      </table>
      <table width="408" border="0" cellspacing="0" cellpadding="0" >
        <tr> 
          <td colspan="2" background="pic/table.gif" height="20">　　 <b>网 站 
            建 设</b>　 <a href="class.asp?classid=1"><img src="pic/more.gif" width="55" height="11" border="0"></a></td>
          <td background="pic/table_r3_c1.gif" rowspan="2">&nbsp;</td>
        </tr>
        <tr> 
          <td background="pic/table_r3_c1.gif" >&nbsp;</td>
          <td width="98%" > 
            <table cellspacing="0" cellpadding="0" width="408" border="0">
              <%    
	sql="select top 5 "    
	sql=sql&"download.id,showname,bb,dateandtime,hits,download.classid,download.Nclassid,Nclass.Nclass"    
	sql=sql&" from download,Nclass where download.classid=1 and download.Nclassid=Nclass.Nclassid "    
	sql=sql&" order by download.dateandtime desc"    
	rs.open sql,conn,1,1    
	if rs.eof and rs.bof then    
%>
              <tr> 
                <td width="75">任何软件</td>
              </tr>
              <%else%>
              <%do while not rs.eof%>
              <tr align="center"> 
                <td width="75" align="left"> <img src="pic/lits.gif" width="13" height="14">&nbsp;<a href="class.asp?classid=<%=rs("classid")%>&Nclassid=<%=rs("Nclassid")%>" style="color: #000000; text-decoration: none"><%=rs("Nclass")%></a></td>
                <td width="345" align="left"> <a href="software.asp?id=<%=rs("id")%>"><%=rs("showname")%></a>( 
                  <%if (month(rs("dateandtime"))&day(rs("dateandtime")))=(month(date())&day(date())) then%>
                  <img src="images/new43.gif" width="29" height="9" border="0"> 
                  <%else%>
                  <%end if%>
                  <%=rs("hits")%>)</td>
                <%      
	rs.movenext          
	loop                     
	end if                    
	rs.close                     
%>
            </table>
          </td>
        </tr>
        <tr> 
          <td colspan="3"><img src="pic/table_r5_c2.gif" width="417" height="26"></td>
        </tr>
      </table>
      <br>
      <table width="408" border="0" cellspacing="0" cellpadding="0" >
        <tr> 
          <td colspan="2" background="pic/table.gif" height="20">　　 <b>Flash 
            相关</b>　 <a href="class.asp?classid=3"><img src="pic/more.gif" width="55" height="11" border="0"></a></td>
          <td background="pic/table_r3_c1.gif" rowspan="2">&nbsp;</td>
        </tr>
        <tr> 
          <td background="pic/table_r3_c1.gif" >&nbsp;</td>
          <td width="98%" > 
            <table cellspacing="0" cellpadding="0" width="408" border="0">
              <%    
	sql="select top 5 "    
	sql=sql&"download.id,showname,bb,dateandtime,hits,download.classid,download.Nclassid,Nclass.Nclass"    
	sql=sql&" from download,Nclass where download.classid=3 and download.Nclassid=Nclass.Nclassid "    
	sql=sql&" order by download.dateandtime desc"    
	rs.open sql,conn,1,1    
	if rs.eof and rs.bof then    
%>
              <tr> 
                <td width="75">任何软件</td>
              </tr>
              <%else%>
              <%do while not rs.eof%>
              <tr align="center"> 
                <td width="75" align="left"><img src="pic/lits.gif" width="13" height="14"><a href="class.asp?classid=<%=rs("classid")%>&Nclassid=<%=rs("Nclassid")%>" style="color: #000000; text-decoration: none"> 
                  <%=rs("Nclass")%></a> </td>
                <td width="345" align="left">&nbsp;<a href="software.asp?id=<%=rs("id")%>"><%=rs("showname")%></a>( 
                  <%if (month(rs("dateandtime"))&day(rs("dateandtime")))=(month(date())&day(date())) then%>
                  <img src="images/new43.gif" width="29" height="9" border="0"> 
                  <%else%>
                  <%end if%>
                  <%=rs("hits")%>)</td>
                <%      
	rs.movenext          
	loop                     
	end if                    
	rs.close                     
%>
            </table>
          </td>
        </tr>
        <tr> 
          <td colspan="3"><img src="pic/table_r5_c2.gif" width="417" height="26"></td>
        </tr>
      </table>
      <br>
      <table width="408" border="0" cellspacing="0" cellpadding="0" >
        <tr> 
          <td colspan="2" background="pic/table.gif" height="20">　　 <b>图 象 
            处 理</b>　 <a href="class.asp?classid=4"><img src="pic/more.gif" width="55" height="11" border="0"></a></td>
          <td background="pic/table_r3_c1.gif" rowspan="2">&nbsp;</td>
        </tr>
        <tr> 
          <td background="pic/table_r3_c1.gif" >&nbsp;</td>
          <td width="98%" > 
            <table cellspacing="0" cellpadding="0" width="408" border="0">
              <%    
	sql="select top 5 "    
	sql=sql&"download.id,showname,bb,dateandtime,hits,download.classid,download.Nclassid,Nclass.Nclass"    
	sql=sql&" from download,Nclass where download.classid=4 and download.Nclassid=Nclass.Nclassid "    
	sql=sql&" order by download.dateandtime desc"    
	rs.open sql,conn,1,1    
	if rs.eof and rs.bof then    
%>
              <tr> 
                <td width="75">任何软件</td>
              </tr>
              <%else%>
              <%do while not rs.eof%>
              <tr align="center"> 
                <td width="75" align="left"><img src="pic/lits.gif" width="13" height="14"><a href="class.asp?classid=<%=rs("classid")%>&Nclassid=<%=rs("Nclassid")%>" style="color: #000000; text-decoration: none"> 
                  <%=rs("Nclass")%></a> </td>
                <td width="345" align="left">&nbsp;<a href="software.asp?id=<%=rs("id")%>"><%=rs("showname")%></a>( 
                  <%if (month(rs("dateandtime"))&day(rs("dateandtime")))=(month(date())&day(date())) then%>
                  <img src="images/new43.gif" width="29" height="9" border="0"> 
                  <%else%>
                  <%end if%>
                  <%=rs("hits")%>)</td>
                <%      
	rs.movenext          
	loop                     
	end if                    
	rs.close                     
%>
            </table>
          </td>
        </tr>
        <tr> 
          <td colspan="3"><img src="pic/table_r5_c2.gif" width="417" height="26"></td>
        </tr>
      </table>
      <br>
      <table width="408" border="0" cellspacing="0" cellpadding="0" >
        <tr> 
          <td colspan="2" background="pic/table.gif" height="20">　　 <b>辅 助 
            工 具</b>　 <a href="class.asp?classid=5"><img src="pic/more.gif" width="55" height="11" border="0"></a></td>
          <td background="pic/table_r3_c1.gif" rowspan="2">&nbsp;</td>
        </tr>
        <tr> 
          <td background="pic/table_r3_c1.gif" >&nbsp;</td>
          <td width="98%" > 
            <table cellspacing="0" cellpadding="0" width="408" border="0">
              <%    
	sql="select top 5 "    
	sql=sql&"download.id,showname,bb,dateandtime,hits,download.classid,download.Nclassid,Nclass.Nclass"    
	sql=sql&" from download,Nclass where download.classid=5 and download.Nclassid=Nclass.Nclassid "    
	sql=sql&" order by download.dateandtime desc"    
	rs.open sql,conn,1,1    
	if rs.eof and rs.bof then    
%>
              <tr> 
                <td width="75">任何软件</td>
              </tr>
              <%else%>
              <%do while not rs.eof%>
              <tr align="center"> 
                <td width="75" align="left"> <img src="pic/lits.gif" width="13" height="14">&nbsp;<a href="class.asp?classid=<%=rs("classid")%>&Nclassid=<%=rs("Nclassid")%>" style="color: #000000; text-decoration: none"><%=rs("Nclass")%></a></td>
                <td width="345" align="left"> <a href="software.asp?id=<%=rs("id")%>"><%=rs("showname")%></a>( 
                  <%if (month(rs("dateandtime"))&day(rs("dateandtime")))=(month(date())&day(date())) then%>
                  <img src="images/new43.gif" width="29" height="9" border="0"> 
                  <%else%>
                  <%end if%>
                  <%=rs("hits")%>)</td>
                <%      
	rs.movenext          
	loop                     
	end if                    
	rs.close                     
%>
            </table>
          </td>
        </tr>
        <tr> 
          <td colspan="3"><img src="pic/table_r5_c2.gif" width="417" height="26"></td>
        </tr>
      </table>
    </TD>
    <TD vAlign="top" width="173" align="right"> <br>
      <table width="168"  border=0 cellpadding="0" cellspacing="0">
        <tr>
          <td class="li" height="22"><strong>网 站 建 设</strong> 分 类</td>
        </tr>
        <tr>
          <td class="3b" background="pic/a-27.gif"><%        
	if request("classid")="" then        
	sql="select Nclass,Nclassid from Nclass where classid=1"        
	else        
	sql="select Nclass,Nclassid from Nclass where classid="&request("classid")        
	end if        
	rs.open sql,conn,1,1%>
              <table width="100%" border="0" cellspacing="0"  cellpadding="0">
                <%do while not rs.eof%>
                <tr>
                  <td align="right" width="10%"  >·</td>
                  <td width="90%"><%if not Rs.eof then%>
                      <span style="letter-spacing: 2pt"> <a href="class.asp?classid=<%=request("classid")%>&Nclassid=<%=rs("Nclassid")%>"><%=Rs("Nclass")%></a></span> </td>
                  <%rs.movenext  
		end if %>
                </tr>
                <%loop        
  Rs.Close        
   %>
            </table></td>
        </tr>
      </table>
      <br>
      <table width="168"  border=0 cellpadding="0" cellspacing="0">
        <tr>
          <td class="li" height="22"><b>Flash 相 关</b> 分 类</td>
        </tr>
        <tr>
          <td class="3b" background="pic/a-27.gif"><%        
	if request("classid")="" then        
	sql="select Nclass,Nclassid from Nclass where classid=3"        
	else        
	sql="select Nclass,Nclassid from Nclass where classid="&request("classid")        
	end if        
	rs.open sql,conn,1,1%>
              <table width="100%" border="0" cellspacing="0"  cellpadding="0">
                <%do while not rs.eof%>
                <tr>
                  <td align="right" width="10%"  >·</td>
                  <td width="90%"><%if not Rs.eof then%>
                      <span style="letter-spacing: 2pt"> <a href="class.asp?classid=<%=request("classid")%>&Nclassid=<%=rs("Nclassid")%>"><%=Rs("Nclass")%></a></span> </td>
                  <%rs.movenext  
		end if %>
                </tr>
                <%loop        
  Rs.Close        
   %>
            </table></td>
        </tr>
      </table>
      <br>
      <table width="168"  border=0 cellpadding="0" cellspacing="0">
        <tr>
          <td class="li" height="22"><b>图 象 处 理</b> 分 类</td>
        </tr>
        <tr>
          <td class="3b" background="pic/a-27.gif"><%        
	if request("classid")="" then        
	sql="select Nclass,Nclassid from Nclass where classid=4"        
	else        
	sql="select Nclass,Nclassid from Nclass where classid="&request("classid")        
	end if        
	rs.open sql,conn,1,1%>
              <table width="100%" border="0" cellspacing="0"  cellpadding="0">
                <%do while not rs.eof%>
                <tr>
                  <td align="right" width="10%"  >·</td>
                  <td width="90%"><%if not Rs.eof then%>
                      <span style="letter-spacing: 2pt"> <a href="class.asp?classid=<%=request("classid")%>&Nclassid=<%=rs("Nclassid")%>"><%=Rs("Nclass")%></a></span> </td>
                  <%rs.movenext  
		end if %>
                </tr>
                <%loop        
  Rs.Close        
   %>
            </table></td>
        </tr>
      </table>
    <br>
    <table width="168"  border=0 cellpadding="0" cellspacing="0">
      <tr>
        <td class="li" height="22"><b>辅 助 工 具</b> 分 类</td>
      </tr>
      <tr>
        <td class="3b" background="pic/a-27.gif"><%        
	if request("classid")="" then        
	sql="select Nclass,Nclassid from Nclass where classid=5"        
	else        
	sql="select Nclass,Nclassid from Nclass where classid="&request("classid")        
	end if        
	rs.open sql,conn,1,1%>
            <table width="100%" border="0" cellspacing="0"  cellpadding="0">
              <%do while not rs.eof%>
              <tr>
                <td align="right" width="10%"  >·</td>
                <td width="90%"><%if not Rs.eof then%>
                    <span style="letter-spacing: 2pt"> <a href="class.asp?classid=<%=request("classid")%>&Nclassid=<%=rs("Nclassid")%>"><%=Rs("Nclass")%></a></span> </td>
                <%rs.movenext  
		end if %>
              </tr>
              <%loop        
  Rs.Close        
   %>
          </table></td>
      </tr>
    </table></TD>
  </tr>
</table>
<tr> 
  <td width="100%" height="1"></td>
</tr>
<tr> 
  <td width="100%" height="5" align="center" valign="top"> </td>
</tr>
<tr> 
  <td width="100%" height="27" align="center" valign="top"> 
</TR>

<br>
<!--#include file="foot.asp"-->
<%   
set rs=nothing    
conn.close    
set conn=nothing%>
</HTML>