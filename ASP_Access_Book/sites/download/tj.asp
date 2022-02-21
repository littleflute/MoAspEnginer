<%@ CODEPAGE = "936" %>
<!--#include file="conn.asp"-->
<%
  dim rs
  set rs = server.createobject("adodb.recordset")
%>
<html>
<head>
<title>推荐软件榜</title>
<META content="text/html; charset=gb2312" http-equiv="Content-Type">
</head>
<!--#include file="head.asp"-->
<!--#include file="lanmu.asp"-->
<table  width="770" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="AutoNumber1">
  <tr> 
    <td height="36" colspan="3" bgcolor="#FFFFFF"> 
      <div align=center><b>推荐软件列表</b></div></td>
  </tr>
  <tr> 
    <td width="384" align="center" valign=top bgcolor="#FFFFFF"> <center> 
      <TABLE width=357 height="142" border=1 cellpadding=0 cellspacing=0  bordercolor=#84b6ad  bordercolordark=#ffffff>
          <TR align=middle> 
            <TD valign="top" width="234">
                  <TABLE cellpadding=4 cellspacing=0 width=350 height="100"  bgcolor="#FFFFFF">
                    <tr> 
                      <td width="313" align="center" background="pic/bar.gif"> 
                        <b>推 荐 软 件</b> </td>
                    </tr>
                    <tr> 
                      <td align="left" width="313"> 
                        <%    
	sql="select top 35 "    
	sql=sql&"download.id,showname,bb,dateandtime,hits,download.classid,download.Nclassid,Nclass.Nclass "    
	sql=sql&" from download,Nclass where download.hots=1 and download.Nclassid=Nclass.Nclassid "    
	sql=sql&" order by download.dateandtime desc"    
	rs.open sql,conn,1,1    
	if rs.eof and rs.bof then    
%>
                        　 
                    <tr> 
                      <td>没有任何软件</td>
                    </tr>
                    <%else%>
                    <%do while not rs.eof%>
                    <tr> 
                      <td width="100%" onMouseOver="this.bgColor='#cccccc';" onMouseOut="this.bgColor='#ffffff';" height="23">&nbsp; 
                        <img src="pic/lits.gif" width="13" height="14">&nbsp;<a href='class.asp?classid=<%=rs("classid")%>&Nclassid=<%=rs("Nclassid")%>'><%=rs("Nclass")%></a>&nbsp;-&nbsp;<a href="software.asp?id=<%=rs("id")%>"><%=rs("showname")%></a><font color='#999999'>( 
                        <%if (month(rs("dateandtime"))&day(rs("dateandtime")))=(month(date())&day(date())) then%>
                        <font color=red> 
                        <%else%>
                        <font color='#999999'> 
                        <%end if%>
                        <%=month(rs("dateandtime"))%>月<%=day(rs("dateandtime"))%>日</font>,<font color=green><%=rs("hits")*3%></font>)</font></font></td>
                      <%      
	rs.movenext          
	loop                     
	end if                    
	rs.close                     
%>
                  </table>
                </div>
          </td>
          </tr>
        </table>
      </center>
      <td width="380" bgcolor="#FFFFFF" valign=top> 
      <table width=235 height="141" border=1 align="center" cellpadding=0 cellspacing=0  bordercolor=#84b6ad bordercolordark=#ffffff>
        <tr align=middle> 
          <td valign="top" width="234">
                <table cellpadding=4 cellspacing=0 width=350 height="100"  bgcolor="#FFFFFF">
                  <tr> 
                    <td width="339" align="center" background="pic/bar.gif"> 
                      <b>精 品 软 件</b> </td>
                  </tr>
                  <tr> 
                    <td align="left" width="339"> 
                      <%    
	sql="select top 35 "    
	sql=sql&"download.id,showname,bb,dateandtime,hits,download.classid,download.Nclassid,Nclass.Nclass "    
	sql=sql&" from download,Nclass where download.stop=1 and download.Nclassid=Nclass.Nclassid "    
	sql=sql&" order by download.dateandtime desc"    
	rs.open sql,conn,1,1    
	if rs.eof and rs.bof then    
%>
                      　 
                  <tr> 
                    <td>没有任何软件</td>
                  </tr>
                  <%else%>
                  <%do while not rs.eof%>
                  <tr> 
                    <td width="100%" onMouseOver="this.bgColor='#cccccc';" onMouseOut="this.bgColor='#ffffff';" height="23">&nbsp; 
                      <img src="pic/lits.gif" width="13" height="14">&nbsp;<a href='class.asp?classid=<%=rs("classid")%>&Nclassid=<%=rs("Nclassid")%>'><%=rs("Nclass")%></a>&nbsp;-&nbsp;<a href="software.asp?id=<%=rs("id")%>"><%=rs("showname")%></a><font color='#999999'>( 
                      <%if (month(rs("dateandtime"))&day(rs("dateandtime")))=(month(date())&day(date())) then%>
                      <font color=red> 
                      <%else%>
                      <font color='#999999'> 
                      <%end if%>
                      <%=month(rs("dateandtime"))%>月<%=day(rs("dateandtime"))%>日</font>,<font color=green><%=rs("hits")*3%></font>)</font></font></td>
                    <%      
	rs.movenext          
	loop                     
	end if                    
	rs.close                     
%>
                </table>
              </center>
            </div></td>
        </tr>
      </table> 
</TABLE>
  
 
<br>
<!--#include file="foot.asp"-->
</body>
</html>