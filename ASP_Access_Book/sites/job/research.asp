<% Response.Buffer=True %> 
<!--#include file="inc/dbconn.inc"-->

<% if request("stype")="" then
options=request("options")
if Request.ServerVariables("REMOTE_ADDR")=request.cookies("IPAddress") then
response.write"<SCRIPT language=JavaScript>alert('��л����֧�֣����Ѿ�Ͷ��Ʊ�ˣ������ظ�ͶƱ��лл��');"
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
response.write"<SCRIPT language=JavaScript>alert('Ŀǰ�����˲�����飡');"
response.write"javascript:window.close();</SCRIPT>"  
end if
total=rs("selecta")+rs("selectb")+rs("selectc")+rs("selectd")   
selecta=(rs("selecta")/total)*100   
selectb=(rs("selectb")/total)*100   
selectc=(rs("selectc")/total)*100   
selectd=(rs("selectd")/total)*100 %><head>
<title>�����˲�==&gt;������</title>
<LINK rel="stylesheet" href="inc/index.css" type="text/css">
</head>

<p align="center">��


<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="366" height="48">
    <tr>
      <td width="366" height="48" valign="top"><font color="#000000">�����˲�==&gt;�˲��г�==&gt;������<br>
        ===============================================<br>
        </font>     
        <font color="#000073">     
                ����Ϊ����Ŀ��Щ���漱��Ľ�?</font> 
    </td>  
    </tr>  
    <tr>
      <td width="366" height="111" valign="top">
        A.���Ի����:&nbsp; <img src=images/research.gif width=<%=int(selecta*2)%> height=8> <%=rs("selecta")%>�� <%=round(selecta,3)%>%<br>            
        B.������Ƹ��Ϣ:<img src=images/research.gif width=<%=int(selectb*2)%> height=8> <%=rs("selectb")%>�� <%=round(selectb,3)%>%<br>            
        C.�ḻ��Ŀ����:<img src=images/research.gif width=<%=int(selectc*2)%> height=8> <%=rs("selectc")%>�� <%=round(selectc,3)%>%<br>            
        D.�Ľ�ҳ�����:<img src=images/research.gif width=<%=int(selectd*2)%> height=8> <%=rs("selectd")%>�� <%=round(selectd,3)%>%&nbsp;<font color="#000073"> 
        <br>       
        <br>
        ���С�<%=total%>���˲μӵ���<br>
        ===============================================</font></center> 
    </td>  
    </tr>  
  </table>  
</div>  
<p align="center">��<a href="javascript:window.close()">�رմ���</a>��    
<% rs.close     
set rs=nothing     
conn.close     
set conn=nothing %> 
