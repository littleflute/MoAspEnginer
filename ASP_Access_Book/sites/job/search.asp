<% Response.Buffer=True %>
<!--#include file="inc/dbconn.inc"-->
<% key=trim(request("key"))
   stype=request("stype") 
   gzdd=request("gzdd") %>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" href="inc/index.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>天天人才—&gt;人才市场—&gt;站内搜索</title>
</head>
<SCRIPT LANGUAGE="JavaScript">
<!--//
function check()
{
   if (isNaN(go2to.page.value))
		alert("请正确填写转到页数！");
   else if (go2to.page.value=="") 
	     {
		alert("请输入转到页数！");
		 }
   else
		go2to.submit();
}
//-->
</SCRIPT>
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" height="358">
    <tr>
  </center>
      
      <td width="293" height="284" valign="top">
           <% set rs=server.createobject("adodb.recordset") 
              if not isempty(request("page")) then   
	          pagecount=cint(request("page"))   
              else   
	          pagecount=1   
              end if
              if stype="company" then 
			  if gzdd="noxz" then
              sql="select * from company where job  like '%"&key&"%' and cname like '%"&key&"%' order by id desc" 
              else 
              sql="select * from company where job like '%"&key&"%' and cname like '%"&key&"%' and gzdd='"&gzdd&"' order by id desc"   
              end if
			  else 
			  if gzdd="noxz" then
              sql="select * from person where job  like '%"&key&"%' and iname like '%"&key&"%' order by id desc" 
              else 
              sql="select * from person where job  like '%"&key&"%' and iname like '%"&key&"%' and gzdd='"&gzdd&"' order by id desc"   
              end if
			  end if
			
              rs.open sql,conn,1,1   
              if rs.eof and rs.bof then   
              response.write"<SCRIPT language=JavaScript>alert('对不起，没有符合搜索条件的记录！');"
              response.write"javascript:window.close();</SCRIPT>"  
              end if
              rs.pagesize=10
              if pagecount>rs.pagecount or pagecount<=0 then              
              pagecount=1              
              end if              
              rs.AbsolutePage=pagecount %>   
  
        <div align="left">
         <table border="1" cellpadding="0" cellspacing="0" width="430" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
          <tr>
          <td height="17" colspan="3" valign="bottom" width="426"></td> 
          </tr>
          <tr>
          <% if rs.pagecount=1 then %>
		  <td width="426" height="6" colspan="3" valign="bottom"><font color="#000000"></font><font color="#000000">共有[<font color="#ff0000"><%=rs.recordcount%></font>]条<% if stype="company" then response.write"招聘信息"  else  response.write"求职简历" end if 
		  %>符合条件 以下是[<font color="red">1～<%=rs.recordcount%></font>]条</font></td>              
          <tr>
          <td width="426" height="7" colspan="3" valign="bottom"></td>
          </tr>
         <%else%>
		 <td width="426" height="6" colspan="3" valign="bottom"><font color="#000000"></font><font color="#000000">
		 <% page_start=(pagecount-1)*rs.pagesize
            if pagecount=1 then page_start=1
		    page_end=rs.pagesize*pagecount
		    if pagecount*rs.pagesize=>rs.recordcount then page_end=rs.recordcount end if%>
         共有[<font color="#ff0000"><%=rs.recordcount%></font>]条<% if stype="company" then response.write"招聘信息"  else  response.write"求职简历" end if %>符合条件    
         以下是[<font color="red"><%=page_start%>～<%=page_end%></font>]条</font></td>           
          <tr> 
		  <td width="426" height="6" colspan="3" valign="bottom">
		  
		  <% response.write"<form name=go2to form method=Post action='search.asp?stype="+stype+"&gzdd="+gzdd+"&key="+key+"'>"
		     response.write "<font color='#000064'>◆&nbsp;</font>"                                              
		     if pagecount=1 then                                                        
			 response.write "<font color='#000064'>首页 上一页</font>&nbsp;"
			 else                                                        
             response.write "<a href=search.asp?page=1&stype="+stype+"&gzdd="+gzdd+"&key="+key+"><font color='0000BE'>首页</font></a>&nbsp;" 
	         response.write "<a href=search.asp?page="+cstr(pagecount-1)+"&stype="+stype+"&gzdd="+gzdd+"&key="+key+"><font color='0000BE'>上一页</font></a>&nbsp;"
			 end if                                             
             if rs.PageCount-pagecount<1 then                                                        
             response.write "<font color='#000064'>下一页 尾页</font>"                                                    
			 else                                                        
             response.write "<a href=search.asp?page="+cstr(pagecount+1)+"&stype="+stype+"&gzdd="+gzdd+"&key="+key+"><font color='0000BE'>下一页</font></a>&nbsp;"
			 response.write "<a href=search.asp?page="+cstr(rs.PageCount)+"&stype="+stype+"&gzdd="+gzdd+"&key="+key+"><font color='0000BE'>尾页</font></a>"           
             end if 
			 response.write "<font color='000064'>&nbsp;页次:<font color=blue>"&pagecount&"</font>/"&rs.pagecount&"页</font>" 
			response.write "<font color='000064'> 转到第<input type='text' name='page' size=2 maxLength=3 style='font-size: 9pt; color:#00006A; position: relative; height: 18' value="&PageCount&">页</font>&nbsp;"                               
			response.write "<input class=button type='button' value='确 定' onclick=check() style='font-family: 宋体; font-size: 9pt; color: #000073; position: relative; height: 19'>" %>         
             </td>
           <% end if %>
           </tr>
           <% if stype="company" then %>
          <tr>
      <td height="3" valign="top" colspan="3" bgcolor="#000000">
      </td>
          </tr>
		  <tr>
          <td width="174" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;公司名称</td>
          <td width="174" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;招聘职位</td>
          <td height="18" bgcolor="#EBEEF3" valign="bottom" width="75"><p align="center">发布日期</p></td>
          </tr>
          <% do while not rs.eof %> 
          <tr>
          <td width="174" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<a 
		  href="company.asp?uid=<%=rs("uname")%>" target="_blank"><% if len(rs("cname"))>12 then%><%=Left(rs("cname"),12)%>...<% else%><%=rs("cname")%><%end if%></a></td>
          <td width="174" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<a 
		  href="job.asp?uid=<%=rs("uname")%>" target="_blank"><%=rs("job")%></a></td>
          <td width="74" height="18" bgcolor="#EBEEF3" valign="bottom"><p align="center"><%=rs("idate")%></p></td>
          </tr>
		 <% i=i+1                                                                                                  
          rs.movenext                                                                                                  
          if i>=rs.PageSize then exit do 
		  loop                                                                    
          rs.close                                                                                                
          set rs=nothing                                                                                                
          conn.close                                                                                                
          set conn=nothing %>
          <tr>
      <td height="3" valign="top" colspan="3" bgcolor="#000000">
      </td>
          </tr>
          <tr>
          <td width="422" height="8" valign="bottom" colspan="3">
            <p align="center"><br>
            【<a href="javascript:window.close()">关闭窗口</a>】</td>
          </tr>
		  </table>
          </div>
		  <% else %>
           <div align="left">
		  <table border="1" cellpadding="0" cellspacing="0" width="430" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
        <tr>
      <td height="3" valign="top" colspan="5" bgcolor="#000000">
      </td>
        </tr>
        <tr>
          <td width="67" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="left">&nbsp;姓名</p>   
          </td>
          <td width="45" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">性 别</p>  
          </td>
          <td width="56" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">学 历   
          </td>
          <td width="178" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="left">&nbsp;应聘职位</td>
  <center>
          <td width="75" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">登录日期</p>
  </td>
        </tr>
        <% do while not rs.eof %>
        </center>
        <tr>
          <td width="64" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="left">&nbsp;<a href="person.asp?uid=<%=rs("uname")%>" target="_blank"><%=rs("iname")%></a></p>
          </td>
          <td width="48" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">[<%=rs("sex")%>]</p>
          </td>
          <center>
          <td width="55" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">[<%=rs("edu")%>]</td>
          <td width="179" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<%=rs("job")%></td>
          <td width="72" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center"><%=rs("idate")%></p>
          </td>
        </tr>
          <% i=i+1                                                                                                  
          rs.movenext                                                                                                  
          if i>=rs.PageSize then exit do 
		  loop                                                                    
          rs.close                                                                                                
          set rs=nothing                                                                                                
          conn.close                                                                                                
          set conn=nothing %> 
          <tr>
      <td height="3" valign="top" colspan="5" bgcolor="#000000">
      </td>
          </tr>
          <tr>
          <td width="418" height="7" valign="bottom" colspan="5">
            <p align="center"><br>
            【<a href="javascript:window.close()">关闭窗口</a>】
          </td>
          </tr>
		 </table>
        </center>
        </div>
		<% end if %>
      </td>
  <center>
    </tr>
  </center>
    </table>
</body>
</html>

















