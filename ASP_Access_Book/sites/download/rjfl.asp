<%@ LANGUAGE = VBScript.Encode %> 
<!--#include file="conn.asp"-->
<!--#include file="home1.asp"-->
<%
  dim rs
  set rs = server.createobject("adodb.recordset")
%>
<HTML>
<HEAD>
<TITLE>软件分类</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
</HEAD>
<!--#include file="head.asp"-->
<!--#include file="lanmu.asp"-->
<br>
<table width="770" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr> 
    <td width="76%" valign="top" rowspan="2" ><br>
      <table width="99%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td> 
            <%Set Nclasslie =Server.CreateObject("ADODB.Recordset") 
sql="select class,classid from class"        
Set classlie = Server.CreateObject("ADODB.Recordset")                          
classlie.open sql,conn,1,1
if not classlie.eof then
do while not classlie.eof
classid=classlie("classid")%>
            <table width="100%" border="0" cellspacing="1" cellpadding="5" align="center" bgcolor="#DCF3DC">
              <tr> 
                <td bgcolor="DCF3DC" colspan="6"> <a href=class.asp?classid=<%=classid%>><u><b><FONT COLOR="008000"><%=classlie("class")%></FONT></b></u></a></td>
              </tr>
              <%sql="select classid,Nclassid,Nclass from Nclass where classid="&classid
		Nclasslie.open sql,conn,1,1
		do while not Nclasslie.eof%>
              <tr> 
                <td align="center" bgcolor="#ffffff" width="17%"> 
                  <%if not Nclasslie.eof then%>
                  <a href=class.asp?classid=<%=classid%>&Nclassid=<%=Nclasslie("Nclassid")%>><font color="#999999"><%=Nclasslie("Nclass")%></font></a> 
                  <%Nclasslie.movenext             
end if%>
                </td>
                <td align="center" bgcolor="#ffffff" width="17%"> 
                  <%if not Nclasslie.eof then%>
                  <a href=class.asp?classid=<%=classid%>&Nclassid=<%=Nclasslie("Nclassid")%>><font color="#999999"><%=Nclasslie("Nclass")%></font></a> 
                  <%Nclasslie.movenext             
end if%>
                </td>
                <td align="center" bgcolor="#ffffff" width="17%"> 
                  <%if not Nclasslie.eof then %>
                  <a href=class.asp?classid=<%=classid%>&Nclassid=<%=Nclasslie("Nclassid")%>><font color="#999999"><%=Nclasslie("Nclass")%></font></a> 
                  <%Nclasslie.movenext             
end if%>
                </td>
                <td align="center" bgcolor="#ffffff" width="17%"> 
                  <%if not Nclasslie.eof then %>
                  <a href=class.asp?classid=<%=classid%>&Nclassid=<%=Nclasslie("Nclassid")%>><font color="#999999"><%=Nclasslie("Nclass")%></font></a> 
                  <%Nclasslie.movenext             
end if%>
                </td>
                <td align="center" bgcolor="#ffffff" width="17%"> 
                  <%if not Nclasslie.eof then %>
                  <a href=class.asp?classid=<%=classid%>&Nclassid=<%=Nclasslie("Nclassid")%>><font color="#999999"><%=Nclasslie("Nclass")%></font></a> 
                  <%Nclasslie.movenext             
end if%>
                </td>
                <td align="center" bgcolor="#ffffff" width="17%"> 
                  <%if not Nclasslie.eof then %>
                  <a href=class.asp?classid=<%=classid%>&Nclassid=<%=Nclasslie("Nclassid")%>><font color="#999999"><%=Nclasslie("Nclass")%></font></a> 
                  <%Nclasslie.movenext             
end if%>
                </td>
              </tr>
              <%loop             
Nclasslie.Close             
%>
            </table>
            <%classlie.movenext             
loop
end if
classlie.close
rs1.close
conn.close
set rs1=nothing
set conn=nothing
set classlie=nothing
set Nclasslie=nothing
%>
          </td>
        </tr>
      </table>
  </tr>
</table>
<!--#include file="foot.asp"-->
