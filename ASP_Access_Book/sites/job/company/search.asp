<% Response.Buffer=True %>
<!--#include file="../inc/company.inc"-->
<% key=trim(request("key")) %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/index.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>天天人才—&gt;首页</title>
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
<!--#include file="../inc/top3.asp"--> 
<div align="center" >
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="358" align="center">
    <tr>
      <td width="778" height="18" valign="top" colspan="5" bgcolor="#53BEB0"> 
        　 </td>
    </tr>
    <tr>
      <td width="139" valign="top" bgcolor="#53BEB0" height="176"> 　 
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="118" height="280">
            <tr>
              <td width="100%" height="163" background="../images/stat-bg.GIF" valign="top">
        <p align="center"><br>
        <a href="main.asp">登录首页</a><br>
        <br>
		<a href="register.asp">登录/更新公司资料</a><br>
        <br>
		<a href="publish.asp">发布/更新招聘信息</a><br>
        <br>
        <a href="../changepwd.asp?stype=company" target="_blank">修改登录密码</a><br>
        <br>
        <a href="search.asp">全部人才列表</a>
              </td>
            </tr>
            <tr>
              <td width="100%" height="117" background="../images/stat-bg.GIF" valign="top">
        <p align="center"><br>
        <br>
        <a href="favorite.asp">我的收藏夹<br>
        <br>
        <a href="mailbox.asp">我的信箱</a><br>
        <br>
        <a href="../exit.asp">退出登录</a>
              </td>
            </tr>
          </table>
        </div>
        <p align="center">&nbsp; </p>
        <p align="center">　
      </td>
      <td width="24" height="284" valign="top"><img border="0" src="../images/selfk.GIF"></td>
  </center>
      
      <td width="480" height="284" valign="top">
           <% set rs=server.createobject("adodb.recordset") 
              if not isempty(request("page")) then   
	          pagecount=cint(request("page"))   
              else   
	          pagecount=1   
              end if
			  sql="select * from person where job  like '%"&key&"%'  order by id desc"
			  

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
  
        <div align="left">
         <table border="1" cellpadding="0" cellspacing="0" width="465" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
          <tr>
          <td height="1" colspan="7" valign="bottom" width="461"><font color="#000000">&nbsp;</font></td> 
          </tr>
          <tr>
          <% if rs.pagecount=1 then %>
		  <td width="461" height="6" colspan="7" valign="bottom"><font color="#000000">&nbsp;</font>
		  <font color="#000000">共有[<font color="#ff0000"><%=rs.recordcount%></font>]条招聘信息符合条件            
		  以下是[<font color="red">1～<%=rs.recordcount%></font>]条</font></td>             
          <tr>
          <td width="461" height="7" colspan="7" valign="bottom"></td>
          </tr>
         <%else%>
		 <td width="461" height="6" colspan="7" valign="bottom"><font color="#000000">&nbsp;</font><font color="#000000">
		 <% page_start=(pagecount-1)*rs.pagesize
            if pagecount=1 then page_start=1
		    page_end=rs.pagesize*pagecount
		    if pagecount*rs.pagesize=>rs.recordcount then page_end=rs.recordcount end if%>
         共有[<font color="#ff0000"><%=rs.recordcount%></font>]条招聘信息符合条件              
         以下是[<font color="red"><%=page_start%>～<%=page_end%></font>]条</font></td>          
          <tr> 
		  <td width="461" height="6" colspan="7" valign="bottom">
		  
		  <% response.write"<form name=go2to form method=Post action='search.asp?key="+key+"'>"
		     response.write "<font color='#000064'>◆&nbsp;</font>"                                              
		     if pagecount=1 then                                                        
			 response.write "<font color='#000064'>首页 上一页</font>&nbsp;"
			 else                                                        
             response.write "<a href=search.asp?page=1&key="+key+"><font color='0000BE'>首页</font></a>&nbsp;" 
	         response.write "<a href=search.asp?page="+cstr(pagecount-1)+"&key="+key+"><font color='0000BE'>上一页</font></a>&nbsp;"
			 end if                                             
             if rs.PageCount-pagecount<1 then                                                        
             response.write "<font color='#000064'>下一页 尾页</font>"                                                    
			 else                                                        
             response.write "<a href=search.asp?page="+cstr(pagecount+1)+"&key="+key+"><font color='0000BE'>下一页</font></a>&nbsp;"
			 response.write "<a href=search.asp?page="+cstr(rs.PageCount)+"&key="+key+"><font color='0000BE'>尾页</font></a>"           
             end if 
			 response.write "<font color='000064'>&nbsp;页次:<font color=blue>"&pagecount&"</font>/"&rs.pagecount&"页</font>" 
			response.write "<font color='000064'> 转到第<input type='text' name='page' size=2 maxLength=3 style='font-size: 9pt; color:#00006A; position: relative; height: 18' value="&PageCount&">页</font>&nbsp;"                               
			response.write "<input class=button type='button' value='确 定' onclick=check() style='font-family: 宋体; font-size: 9pt; color: #000073; position: relative; height: 19'>" %>         
             </td>
           <%end if%>
		   </form>
           </tr>         
          <tr>
      <td height="3" valign="top" colspan="7" bgcolor="#000000" width="461">
      </td>
          </tr>
		  <tr>
          <td width="56" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;姓名</td>
          <td width="36" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">性别</td>
          <td width="53" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">学历</td>
          <td width="161" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;应聘职位</td>
          <td height="18" bgcolor="#EBEEF3" valign="bottom" width="77"><p align="center">发布日期</p></td>
          <td height="18" bgcolor="#EBEEF3" valign="bottom" width="36">
            <p align="center">发信</td>
          <td height="18" bgcolor="#EBEEF3" valign="bottom" width="30">
            <p align="center">收藏</td>
          </tr>
          <% do while not rs.eof %> 
          <tr>
          <td width="56" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<a href="javascript:openwin('../person.asp?uid=<%=rs("uname")%>','top=10,left=300,width=460,height=420')"><%=rs("iname")%></a></td>
          <td width="36" height="18" bgcolor="#EBEEF3" valign="bottom"><p align="center">[<%=rs("sex")%>]</p></td>
          <td width="53" height="18" bgcolor="#EBEEF3" valign="bottom"><p align="center">[<%=rs("edu")%>]</p></td>
          <td width="161" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<%=rs("job")%></td>
          <td width="77" height="18" bgcolor="#EBEEF3" valign="bottom"><p align="center">[<%=rs("idate")%>]</p></td>
          <td width="36" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center"><a href="javascript:openwin('sendmail.asp?reid=<%=rs("uname")%>','top=10,left=300,width=460,height=420')"><font face="Wingdings" color="#000000" size="2">*</font></a></td>
          <td width="30" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center"><font face="Wingdings"><a href="addfav.asp?fav=<%=rs("uname")%>">1</a></font></td>
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
      <td height="3" valign="top" colspan="7" bgcolor="#000000" width="461">
      </td>
          </tr>
          <tr>
          <td width="461" height="8" valign="bottom" colspan="7">
            <form action="search.asp" method=post>
            <p align="center"><br>
            <font color="#000000">
                职位搜索器: </font><INPUT           
                  maxLength=20 size=16 name="key" style="background-color: #EBEBEB; color: #00006A; font-family: 宋体; font-size: 9pt" value="请输入关键字-S" onclick="Javascript:this.value='';">&nbsp;        
            <input type="submit" value="搜 索" style="font-family: 宋体; font-size: 9pt; color: #00006A">
      </form>
      <br>
          </td>
          </tr>
		  </table>
          </div>
</td>
  <center>
      <td width="1" height="284" valign="top" bgcolor="#00006A"></td>
      <td width="126" height="284" valign="top" bgcolor="#F3F3F3">　</td>
    </tr>
    <tr>
      <td height="1" valign="top" colspan="5" bgcolor="#53BEB0" width="778"> 
        <p align="center"> </td>
    </tr>
  </center>
    </table>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="1">
    <tr>
      <td width="743" height="3" valign="top">
      <p align="center">
      </td>
    </tr>
    <tr>
      <td width="743" height="1" valign="top">
      <p align="center"><script language="javascript" src="../inc/copyright.js"></script>
      </td>
    </tr>
    <tr>
      <td width="743" height="3" valign="top">
      </td>
    </tr>
  </table>
</div>
</body>
</html>
