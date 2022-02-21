<!--#include file="rscoon.asp"-->
<!--#include file="mzfunc.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>投票结果</title>
<style type=text/css>
<!--
BODY{
	margin:0px;
	FONT-SIZE: 12px;
	FONT-FAMILY:  "宋体", "Verdana", "Arial", "Helvetica", "sans-serif";
	scrollbar-face-color: #DEE3E7;
	scrollbar-highlight-color: #FFFFFF;
	scrollbar-shadow-color: #DEE3E7;
	scrollbar-3dlight-color: #D1D7DC;
	scrollbar-arrow-color:  #006699;
	scrollbar-track-color: #EFEFEF;
	scrollbar-darkshadow-color: #98AAB1;
}
table  { border:0px; }
td  { font:normal 12px 宋体; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px 宋体; color:#000000; text-decoration:none; }
a:hover  { color:#428EFF;text-decoration:underline; }
.sec_menu  { border-left:1px solid white; border-right:1px solid white; border-bottom:1px solid white; overflow:hidden; background:#D6DFF7; }
.menu_title  { }
.menu_title span  { position:relative; top:2px; left:8px; color:#215DC6; font-weight:bold; }
.menu_title2  { }
.menu_title2 span  { position:relative; top:2px; left:8px; color:#428EFF; font-weight:bold; }
input,select,Textarea{
font-family:宋体,Verdana, Arial, Helvetica, sans-serif; font-size: 12px;}
}
-->
</style>
</head>
<body>
  <% 
    dim action,id,sql,rs,statid
  action=request.querystring("action")
      if action="poll" then
	     statid=request.Form("statid")
	   set rs = server.createobject("adodb.recordset")
   sql="select lastip from stat where id="&statid
	     rs.open sql,conn,1,3
		  if getIP()=rs("lastip") then
		 response.write "<script>alert('您已经投过票了，请不要重复投票');history.back();</script>"
	      else
		   rs("lastip")=getip()
		   rs.update
		   rs.close
		   end if
		 id=request.Form("rv")
		 if id="" or isnull(id) or isnumeric(id)<>true then
		 response.write "<script>alert('请选择你所投的对象,不能投空票');history.back();</script>"
	     else
		 sql="select * from votes where voteid="&id
		 rs.open sql,conn,1,3
		 end if
		  rs("counts")=rs("counts")+1
		  rs.update
		 response.write "<script>alert('谢谢您的投票,有了您的支持我们将更努的做好服务');window.open('votes.asp?action=look','','width=400,height=300,top='+(screen.availHeight-240)/2+',left='+(screen.availWidth-400)/2);history.back();</script>"
  end if
       if action="look" then
	   sql="select title,id from stat where isvote='true'"
	    set rs=conn.execute(sql)
		statid=rs("id")
	   dim total
     sql="select sum(counts) as total from votes where fromid="&statid
	  set rs=conn.execute(sql)
	  total=rs("total")
       
 %>
  <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
    <tr > 
      <td height="33" colspan="4" align="center" background="images/dhbg2.gif"> 
        目前共有<font color="#ff0000"><%=total%></font>人参予了投票</td>
  </tr>
  <%
set rs=server.createobject("adodb.recordset")
rs.open "select * from votes where fromid="&statid,conn,1,1
if not rs.eof then
do while not rs.eof
%>
  <tr> 
      <td width="25%" height="20"> <%=rs("votetitle")%> </td>
      <td width="43%" height="20"> <img src="images/vb1.gif" border="0" height="10" width="<%=((rs("counts")/total)*100)%>"> 
      </td>
    <td width="32%" height="20"> <font color="#FCBF50"><%=formatnumber((rs("counts")/total)*100,2)%>%</font> </td>
  </tr>
  <%
rs.movenext
loop
end if
rs.close
response.write"<tr><td height=20 colspan=4 align=center>【谢谢您的投票】<br>"&_
"<a href='javascript:window.close()'>关闭</a></td></tr></table>"
 end if 
response.write"</body></html>"
%>