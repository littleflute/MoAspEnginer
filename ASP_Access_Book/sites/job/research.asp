<% Response.Buffer=True %> 
<!--#include file="inc/dbconn.inc"-->

<% if request("stype")="" then
options=request("options")
if Request.ServerVariables("REMOTE_ADDR")=request.cookies("IPAddress") then
response.write"<SCRIPT language=JavaScript>alert('感谢您的支持，您已经投过票了，请勿重复投票，谢谢！');"
response.write"javascript:window.close();</SCRIPT>" 
end if
response.cookies("IPAddress")=Request.ServerVariables("REMOTE_ADDR") 
set rs=server.createobject("adodb.recordset")
sql1="update research set select"&options&"=select"&options&"+1 where id=1"
rs.open sql1,conn,3,3                                                                   
set rs=nothing 
end if %>
<% set rs=server.createobject("adodb.recordset")
sql2="select * from research where id=1"   
rs.open sql2,conn,1,1
if rs("selecta")="0" and rs("selectb")="0" and rs("selectc")="0" and rs("selectd")="0" then
response.write"<SCRIPT language=JavaScript>alert('目前尚无人参与调查！');"
response.write"javascript:window.close();</SCRIPT>"  
end if
total=rs("selecta")+rs("selectb")+rs("selectc")+rs("selectd")   
selecta=(rs("selecta")/total)*100   
selectb=(rs("selectb")/total)*100   
selectc=(rs("selectc")/total)*100   
selectd=(rs("selectd")/total)*100 %><head>
<title>天天人才==&gt;调查结果</title>
<LINK rel="stylesheet" href="inc/index.css" type="text/css">
</head>

<p align="center">　


<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="366" height="48">
    <tr>
      <td width="366" height="48" valign="top"><font color="#000000">天天人才==&gt;人才市场==&gt;调查结果<br>
        ===============================================<br>
        </font>     
        <font color="#000073">     
                您认为本栏目哪些方面急需改进?</font> 
    </td>  
    </tr>  
    <tr>
      <td width="366" height="111" valign="top">
        A.个性化设计:&nbsp; <img src=images/research.gif width=<%=int(selecta*2)%> height=8> <%=rs("selecta")%>人 <%=round(selecta,3)%>%<br>            
        B.增加招聘信息:<img src=images/research.gif width=<%=int(selectb*2)%> height=8> <%=rs("selectb")%>人 <%=round(selectb,3)%>%<br>            
        C.丰富栏目内容:<img src=images/research.gif width=<%=int(selectc*2)%> height=8> <%=rs("selectc")%>人 <%=round(selectc,3)%>%<br>            
        D.改进页面设计:<img src=images/research.gif width=<%=int(selectd*2)%> height=8> <%=rs("selectd")%>人 <%=round(selectd,3)%>%&nbsp;<font color="#000073"> 
        <br>       
        <br>
        共有【<%=total%>】人参加调查<br>
        ===============================================</font></center> 
    </td>  
    </tr>  
  </table>  
</div>  
<p align="center">【<a href="javascript:window.close()">关闭窗口</a>】    
<% rs.close     
set rs=nothing     
conn.close     
set conn=nothing %> 
