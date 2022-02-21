<% Response.Buffer=True %>
<!--#include file="../inc/admin.inc"-->
<% key=trim(request("key")) 
if request("del")<>"" then 
conn.Execute("delete  from company where uname='"&request("del")&"'")
conn.Execute("delete  from cmailbox where reid='"&request("del")&"'")
conn.Execute("delete  from cfavorite where uname='"&request("del")&"'")
conn.Execute("delete  from pfavorite where fuid='"&request("del")&"'")
end if %>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" href="../inc/index.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>天天人才—&gt;人才市场—&gt;管理单位用户</title>
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
<SCRIPT language=JavaScript src="../inc/window.js"></SCRIPT>
<body topmargin="0" leftmargin="0">

  <table border="0" cellpadding="0" cellspacing="0" width="780" height="450">
    
  <tr bgcolor="#53BEB0"> 
    <td width="778" height="14" valign="top" colspan="3"></td>
    </tr>
  <tr>
    <td valign="top" height="424" width="176" bgcolor="#E1E1E1">
      　
      <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" width="118" background="../images/stat-bg.GIF">
          <tr>
            <td height="34">
      <p align="center">
      <a href="mnews.asp">
      新闻管理</a>
            </td>
          </tr>
          <tr>
            <td height="29">
      <p align="center"><a href="mperson.asp">管理个人用户</a>
            </td>
          </tr>
          <tr>
            <td height="34">
      <p align="center"><a href="mcompany.asp">管理单位用户</a>
            </td>
          </tr>
          <tr>
            <td height="34">
      <p align="center"><a href="#"onclick="javascript:if (confirm('是否确定清空调查引擎数据?')) href='reset.asp'; 
	    else return;">清空调查数据</a>
            </td>
          </tr>
          <tr>
            <td height="28">
      <p align="center"><a href="exit.asp">退出登录</a>
            </td>
          </tr>
        </table>
        </center>
      </div>
      <p align="center">&nbsp; </td>
    <td valign="top" height="424" width="1" bgcolor="#53BEB0"> 　 </td>
     <% set rs=server.createobject("adodb.recordset") 
     if not isempty(request("page")) then   
     pagecount=cint(request("page"))   
     else   
     pagecount=1   
     end if
     sql="select * from company where uname like '%"&key&"%' and cname like '%"&key&"%' order by id desc" 
	 rs.open sql,conn,1,1   
     if rs.eof and rs.bof then   
     response.write"<SCRIPT language=JavaScript>alert('对不起，没有符合搜索条件的记录！');"
     response.write"javascript:history.go(-1)</SCRIPT>"       
     end if
     rs.pagesize=10
     if pagecount>rs.pagecount or pagecount<=0 then              
     pagecount=1              
     end if              
     rs.AbsolutePage=pagecount %>   
	
	<td valign="top" height="424" width="602">
    <div align="center">
      <center>
      <table border="1" cellpadding="0" cellspacing="0" width="545" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
        <tr>
          <td height="1" colspan="5" valign="bottom" width="541"><font color="#000000">&nbsp;</font></td> 
        </tr>
        <tr>
		   <% if rs.pagecount=1 then %>
		   <td width="541" height="6" colspan="5" valign="bottom">
          <font color="#000000">共有[<font color="#ff0000"><%=rs.recordcount%></font>]位单位用户符合条件                         
		  以下是[<font color="red">1～<%=rs.recordcount%></font>]条</font></td>             
        </tr>
        <tr>
          <td width="541" height="7" colspan="5" valign="bottom"></td>
        </tr>
        <%else%>
		<tr>
		 <td width="541" height="6" colspan="5" valign="bottom"><font color="#000000">
		 <% page_start=(pagecount-1)*rs.pagesize
            if pagecount=1 then page_start=1
		    page_end=rs.pagesize*pagecount
		    if pagecount*rs.pagesize=>rs.recordcount then page_end=rs.recordcount end if%>
         共有[<font color="#ff0000"><%=rs.recordcount%></font>]位单位用户符合条件                           
         以下是[<font color="red"><%=page_start%>～<%=page_end%></font>]条</font></td>          
        </tr>
        <tr>
		  <td width="541" height="6" colspan="5" valign="bottom">
		  <% response.write"<form name=go2to form method=Post action='mcompany.asp?key="+key+"'>"
		     response.write "<font color='#000064'>◆&nbsp;</font>"                                              
		     if pagecount=1 then                                                        
			 response.write "<font color='#000064'>首页 上一页</font>&nbsp;"
			 else                                                        
             response.write "<a href=mcompany.asp?page=1&key="+key+"><font color='0000BE'>首页</font></a>&nbsp;" 
	         response.write "<a href=mcompany.asp?page="+cstr(pagecount-1)+"&key="+key+"><font color='0000BE'>上一页</font></a>&nbsp;"
			 end if                                             
             if rs.PageCount-pagecount<1 then                                                        
             response.write "<font color='#000064'>下一页 尾页</font>"                                                    
			 else                                                        
             response.write "<a href=mcompany.asp?page="+cstr(pagecount+1)+"&key="+key+"><font color='0000BE'>下一页</font></a>&nbsp;"
			 response.write "<a href=mcompany.asp?page="+cstr(rs.PageCount)+"&key="+key+"><font color='0000BE'>尾页</font></a>"           
             end if 
			 response.write "<font color='000064'>&nbsp;页次:<font color=blue>"&pagecount&"</font>/"&rs.pagecount&"页</font>" 
			response.write "<font color='000064'> 转到第<input type='text' name='page' size=2 maxLength=3 style='font-size: 9pt; color:#00006A; position: relative; height: 18' value="&PageCount&">页</font>&nbsp;"                               
			response.write "<input class=button type='button' value='确 定' onclick=check() style='font-family: 宋体; font-size: 9pt; color: #000073; position: relative; height: 19'>" %>         
             </td>
        </tr>
      </form>
	  <%end if%>
	  <tr>
      <td height="3" valign="top" colspan="5" bgcolor="#000000" width="541">
      </td>
        </tr>
        <tr>
          <td width="75" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;UserID</td>
          <td width="150" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;公司名称</td>
          <td width="163" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;招聘职位</td>
          <td height="18" bgcolor="#EBEEF3" valign="bottom" width="84"><p align="center">发布日期</p></td>
          <td height="18" bgcolor="#EBEEF3" valign="bottom" width="61">
            <p align="center">- 删除 -</td>      
        </tr>
         <% do while not rs.eof %>
         <tr>
          <td width="75" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<%=rs("uname")%></td>
          <td width="150" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<% if rs("cname")<>"" then%><a href="javascript:openwin('../company.asp?uid=<%=rs("uname")%>','top=10,left=300,width=460,height=420')"><% if len(rs("cname"))>10 then%><%=Left(rs("cname"),10)%>...<%else%><%=rs("cname")%><%end if%></a><%else response.write"[未登录]" end if%> 
		  </td>
          <td width="163" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<% if rs("job")<>"" then%><a 
		 href="javascript:openwin('../job.asp?uid=<%=rs("uname")%>','top=10,left=300,width=460,height=420')"><%=rs("job")%>
		 </a><%else response.write"[未登录]" end if%></td>
          <td width="84" height="18" bgcolor="#EBEEF3" valign="bottom"><p align="center">[<% if rs("idate")<>"" then%><%=rs("idate")%><%else response.write"未登录" end if%>]</p></td>
          <td width="61" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center"><font color="#000046">[</font><a href="mcompany.asp?del=<%=rs("uname")%>&key=<%=key%>"><font color="#000046">删除</font></a><font color="#000046">]</font></td>
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
      <td height="3" valign="top" colspan="5" bgcolor="#000000" width="541">
      </td>
        </tr>
        <tr>
          <td width="541" height="8" valign="bottom" colspan="5">
            <form action="mcompany.asp" method=post>
                  <p align="center"><br>
                    <font color="#000000">单位用户搜索: </font>
                    <INPUT                    
                  maxLength=20 size=16 name="key" style="background-color: #EBEBEB; color: #00006A; font-family: 宋体; font-size: 9pt" value="请输入关键字-S" onclick="Javascript:this.value='';">
                    &nbsp; 
                    <input type="submit" value="搜 索" style="font-family: 宋体; font-size: 9pt; color: #00006A">
                    <br>
                    <br>
                </form>
      <br>
          </td>
        </tr>
      </table>
      </center>
    </div>
    </td>
  </tr>
    
  <tr bgcolor="#53BEB0"> 
    <td valign="top" height="12" width="780" colspan="3"> 　 </td>
    </tr>
  </table>
</body>

</html>