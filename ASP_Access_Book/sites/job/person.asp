<% Response.Buffer=True %>
<!--#include file="inc/dbconn.inc"-->
<% uid=request("uid")
set rs=server.createobject("adodb.recordset")
sql1="update person set click=click+1 where uname='"&uid&"'"
rs.open sql1,conn,1,1
sql2="select * from person where job<>'""' and uname='"&uid&"'"
rs.open sql2,conn,1,1 %>
<% if rs.eof or rs.bof then
   response.write"<SCRIPT language=JavaScript>alert('对不起，该用户不存在或已被删除！');"
   response.write"javascript:window.close();</SCRIPT>" 
   end if %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/index.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>天天人才—&gt;<%=rs("iname")%></title>
</head>
<div align="center">
  <center>
  <table border="2" cellpadding="0" cellspacing="0" width="390" height="18" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
    <tr>
      <td height="19" bgcolor="#C6CEDE" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" valign="bottom" colspan="2" width="384">
        <p align="center">=== 个人求职简历 ===</td>                                     
    </tr>
    <tr>
      <td height="19" valign="bottom" width="176">&nbsp;姓名:<%=rs("cname")%></td>
  </center>
   
      <td height="19" valign="bottom" width="206">
        <p align="right">点击数:<%=rs("click")%>&nbsp;</p>
    </td>
    </tr>
  <center>
    <tr>
      <td height="19" valign="bottom" bgcolor="#EBEEF3" colspan="2" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="384">&nbsp;个人基本资料    
        ------------------------------------------------&gt;</td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;性别:<%=rs("sex")%>&nbsp;出生年月:<%=rs("bday")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">&nbsp;民族:<%=rs("mzhu")%>&nbsp;政治面貌:<%=rs("zzmm")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;户籍所在地:<%=rs("hka")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">&nbsp;最高教育程度:<%=rs("edu")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;专 业:<%=rs("zye")%></td>                     
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">&nbsp;毕业院校:<%=rs("school")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;现有职称:<%=rs("zchen")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#EBEEF3" colspan="2" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="384">&nbsp;希望工作条件—联系信息    
        --------------------------------------&gt;</td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;求职类型:<%=rs("jobtype")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">&nbsp;应聘岗位:<%=rs("job")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;希望工作地点:<%=rs("gzdd")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">&nbsp;薪金要求:<%if rs("yuex")="面议" then response.write"面议" else response.write"月薪["&rs("yuex")&"]RMB" end if%></td>
    </tr>
    <% if rs("otheryq")<>"无其他要求" then%>
	<tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;其他要求:</td>
    </tr>
    <tr>
      <td height="3" valign="top" bgcolor="#00006A" colspan="2" width="384">
      </td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="379" align="right">
            <tr>
              <td width="377">&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("otheryq")%></td>
            </tr>
          </table>
        </div>
        </td>
    </tr>
    <tr>
      <td height="3" valign="top" bgcolor="#00006A" colspan="2" width="384">
      </td>
    </tr>
    <%end if%>
	<tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;联系人:<%=rs("cname")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">&nbsp;联系电话:<%=rs("phone")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;寻呼机号码:<%=rs("callnum")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">&nbsp;E-mail:<a href="mailto:<%=rs("email")%>"><%=rs("email")%></a></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;OICQ号码:<%=rs("oicq")%></td>
    </tr>
    <tr>
  <td height="19" valign="bottom" colspan="2" bgcolor="#F3F3F3" width="384">&nbsp;个人主页:<a href="<%=rs("http")%>"><%=rs("http")%></a></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;联系地址:<%=rs("address")%></td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#EBEEF3" colspan="2" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="384">&nbsp;个人主要特长-相关工作经历    
        -----------------------------------&gt;</td>
    </tr>
    <tr>
      <td height="19" valign="bottom" colspan="2" width="384">&nbsp;外语特长:<%=rs("language")%> 等级:<%=rs("lanlevel")%></td>                 
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">&nbsp;普通话程度:<%=rs("pthua")%>&nbsp;计算机能力:<%=rs("computer")%></td>   
    </tr>
    <% if rs("othertc")<>"无其他特长" then%>
	<tr>
      <td height="19" valign="bottom" bgcolor="#FFFFFF" colspan="2" width="384">&nbsp;其它主要特长:</td>
    </tr>
    <tr>
      <td height="3" valign="top" bgcolor="#00006A" colspan="2" width="384">
      </td>
    </tr>
    <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="379" align="right">
            <tr>
              <td width="377">&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("othertc")%></td>
            </tr>
          </table>
        </div>
        </td>
    </tr>
    <tr>
      <td height="3" valign="top" bgcolor="#00006A" colspan="2" width="384">
      </td>
    </tr>
  </center>
   <%end if%>
  <tr>
      <td height="19" valign="bottom" width="176">&nbsp;详细工作经历:</td>
      <td height="19" valign="bottom" width="206">
        <p align="right">相关工作经验:<%=rs("gznum")%>年&nbsp;&nbsp;</p>
      </td>
  </tr>
  <tr>
      <td height="3" valign="top" bgcolor="#00006A" colspan="2" width="384">
      </td>
  </tr>
  <tr>
      <td height="19" valign="bottom" bgcolor="#F3F3F3" colspan="2" width="384">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="379" align="right">
            <tr>
              <td width="377">&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("gzjl")%></td>
            </tr>
          </table>
        </div>
        </td>
  </tr>
  <tr>
      <td height="3" valign="top" bgcolor="#00006A" colspan="2" width="384">
      </td>
  </tr>
	 <% if session("cuid")<>"" then%>
	 <tr>
      <td height="24" valign="bottom" colspan="2" width="384">
        <p align="center">[<a href="company/addfav.asp?fav=<%=rs("uname")%>">添加到收藏夹</a>]                        
		[<a href="company/sendmail.asp?reid=<%=rs("uname")%>">发送站内信件</a>]</td> 
    </tr>
	<%end if%>
  </table>
</div>
<br>
<CENTER>【<a href="javascript:window.close()">关闭窗口</a>】</CENTER>
</html>

