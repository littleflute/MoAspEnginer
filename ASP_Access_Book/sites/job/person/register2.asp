<% Response.Buffer=True %>
<!--#include file="../inc/person.inc"-->
<!--#include file="../inc/html.inc"-->
<% uname=session("puid") 
   modify=request("modify")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from person where uname='"&uname&"'"
rs.open sql,conn,1,1 
if rs("iname")<>"" then
else 
response.write"<SCRIPT language=JavaScript>alert('用户非法操作，请按照顺序填写求职简历！');"
response.write"javascript:history.go(-1)</SCRIPT>"
end if %>
<% if modify<>"ture" and rs("job") <> "" then
   response.write"<SCRIPT language=JavaScript>alert('你已经登录个人简历，请不要重复登录！');"
   response.write"javascript:history.go(-1)</SCRIPT>"
   end if %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" href="../inc/register.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<% if modify<>"ture" then %>
<title>天天人才—&gt;登录求职简历</title>
<%else%>
<title>天天人才—&gt;更新求职简历</title>
</head>
<%end if%>
<SCRIPT LANGUAGE="JavaScript">
<!--//
function check()
{
   if (document.register.gznum.value.length<1) 
		alert("请输入工作经验年数！");
   else if (isNaN(register.gznum.value))
		alert("工作经验栏只能填写数字！");
   else if (document.register.gzjl.value.length<1) 
		alert("请输入详细工作经历！");
   else
		document.register.submit();
}
//-->
</SCRIPT>
<% if modify<>"ture" then %>
<FORM name=register action=register2.asp method=post>
<%else%>
<FORM name="register" action="register2.asp?modify=ture" method="post">
<%end if%>
<body topmargin="0" leftmargin="0">
<!--#include file="../inc/top2.htm"-->
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="340">
    <tr>
        <td width="778" height="18" valign="top" colspan="5" bgcolor="#53BEB0"> 
          　 </td>
    </tr>
    <tr>
        <td width="139" valign="top" bgcolor="#53BEB0"> 
          <p align="center">　 
          <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="118" height="280">
            <tr>
              <td width="100%" height="163" background="../images/stat-bg.GIF" valign="top">
        <p align="center"><br>
        <a href="main.asp">登录首页</a><br>
        <br>
        <a href="register.asp">登录求职简历</a><br>
        <br>
        <a href="modify.asp">更新求职简历</a><br>
        <br>
        <a href="../changepwd.asp?stype=person" target="_blank">修改登录密码</a><br>
        <br><a href="search.asp">全部职位列表</a>
              </td>
            </tr>
            <tr>
              <td width="100%" height="117" background="../images/stat-bg.GIF" valign="top">
        <p align="center"><br>
        <br><a href="favorite.asp">我的收藏夹</a><br>
        <br>
        <a href="mailbox.asp">我的信箱</a><br>
        <br>
        <a href="../exit.asp">退出登录</a>
              </td>
            </tr>
          </table>
        </div>
          <p align="center">&nbsp; </td>
      <td width="36" height="604" valign="top"><img border="0" src="../images/selfk.GIF"></td>
      <td width="478" height="248" valign="top">
  </center>
    
      　
    
        <table border="1" cellpadding="0" cellspacing="0" width="93%" height="258" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
          <tr>
         <td width="100%" height="18" valign="bottom" bgcolor="#C6CEDE" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
              <p align="center">=== 个人主要特长 ===</td>                                              
          </tr>
          <tr>
            <td width="100%" height="124" valign="top" bgcolor="#F3F3F3">
              <p align="center"> 
              <br>
              外语特长:<SELECT 
                  size=1 name=language> <OPTION value=无 <%if rs("language") ="无" then Response.Write "selected"%>>无
                  <OPTION value=英语 <%if rs("language") ="英语" then Response.Write "selected"%>>英语
                  <OPTION value=日语 <%if rs("language") ="日语" then Response.Write "selected"%>>日语
                  <OPTION value=法语 <%if rs("language") ="法语" then Response.Write "selected"%>>法语
                  <OPTION value=德语 <%if rs("language") ="德语" then Response.Write "selected"%>>德语
                  <OPTION value=俄语 <%if rs("language") ="俄语" then Response.Write "selected"%>>俄语
                  <OPTION value=朝鲜语 <%if rs("language") ="朝鲜语" then Response.Write "selected"%>>朝鲜语
                  <OPTION value=西班牙语 <%if rs("language") ="西班牙语" then Response.Write "selected"%>>西班牙语
                  <OPTION value=阿拉伯语 <%if rs("language") ="阿拉伯语" then Response.Write "selected"%>>阿拉伯语
                  <OPTION value=其它 <%if rs("language") ="其它" then Response.Write "selected"%>>其它</OPTION></SELECT>&nbsp;                                  
                  <SELECT size=1 name=lanlevel> 
                  <OPTION value=无 <%if rs("lanlevel") ="无" then Response.Write "selected"%>>无
                  <OPTION value=四级  <%if rs("lanlevel") ="四级" then Response.Write "selected"%>>四级
                  <OPTION value=八级  <%if rs("lanlevel") ="八级" then Response.Write "selected"%>>八级
                  <OPTION value=六级  <%if rs("lanlevel") ="六级" then Response.Write "selected"%>>六级
                  <OPTION value=熟练  <%if rs("lanlevel") ="熟练" then Response.Write "selected"%>>熟练
                  <OPTION value=精通  <%if rs("lanlevel") ="精通" then Response.Write "selected"%>>精通
                  <OPTION value=良好  <%if rs("lanlevel") ="良好" then Response.Write "selected"%>>良好
                  <OPTION value=一般  
				  <%if rs("lanlevel") ="一般" then Response.Write "selected"%>>一般</OPTION></SELECT></p> 
              <p align="center">普通话程度:<SELECT 
                  size=1 name=pthua> 
                  <OPTION value=标准 <%if rs("pthua") ="标准" then Response.Write "selected"%>>精通
                  <OPTION value=一般 <%if rs("pthua") ="一般" then Response.Write "selected"%>>一般
                  <OPTION value=较差 <%if rs("pthua") ="较差" then Response.Write "selected"%>>较差
                  </OPTION></SELECT>&nbsp; 计算机能力:                                      
                  <SELECT size=1 name=computer> 
                  <OPTION value=一般 <%if rs("computer") ="一般" then Response.Write "selected"%>>一般
                  <OPTION value=优秀 <%if rs("computer") ="优秀" then Response.Write "selected"%>>优秀
                  <OPTION value=良好 <%if rs("computer") ="良好" then Response.Write "selected"%>>良好
                    <OPTION value=较差 <%if rs("computer") ="较差" then Response.Write "selected"%>>较差</OPTION></SELECT></p> 
              <p align="center">其他主要特长:<br>
              <font color="#FF0000">
              （限100字以内）</font><br>
              <%if modify<>"ture" then 
			  kothertc=""
			  else
			  kothertc=replace(rs("othertc"),"<br>",chr(13))
              kothertc=replace(kothertc,"&nbsp;"," ")%>
			  <%end if%>          
              <textarea rows="6" name="othertc" cols="34"><%=kothertc%></textarea>          
              <br>
              <br>
			  </p> 
            </td>
          </tr>
          <tr>
            <td width="100%" height="18" valign="bottom" bgcolor="#C6CEDE" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
              <p align="center">=== 相关工作经历 ===                                     
            </td>
          </tr>
          <tr>
            <td width="100%" height="108" valign="top" bgcolor="#F3F3F3">
              <p align="center"><br>
              工作经验:至今相关工作经验共有<INPUT 
                  maxLength=2 size=2 name=gznum value="<%=rs("gznum")%>">年</p>
              <p align="center">详细工作经历:<BR>按照格式<font color="#FF0000">{始(年.月) 至(年.月) 职务名称 公司名称}</font>填写         
			  <%if modify<>"ture" then 
			  kgzjl=""
			  else
			  kgzjl=replace(rs("gzjl"),"<br>",chr(13))
              kgzjl=replace(kgzjl,"&nbsp;"," ")%>
			  <%end if%>          
              <textarea rows="11" name="gzjl" cols="48"><%=kgzjl%></textarea>          
			  </p> 
              <p align="center">
			  <% if modify<>"ture" then %>
			  <input type="button" value="上一步" onclick="javascript:history.go(-1)">  
              <input type="button" value="下一步" onClick="check()"><br><br>  
              <%else%>
              <input type="button" value="更 新" onClick="check()"> 
			  <br>
			  <br>
			  <%end if%>
            </td>
          </tr>
        </table>
      <% rs.close %>
      </td>
  <center>
      <td width="1" height="604" valign="top" bgcolor="#00006A"></td>
      <td width="119" height="604" valign="top" bgcolor="#F3F3F3">　</td>
    </tr>
    <tr>
        <td width="778" height="1" valign="top" colspan="5" bgcolor="#53BEB0"> 
          <p align="center"> </td>
    </tr>
    <tr>
      <td width="778" height="3" valign="top" colspan="5">
      <p align="center">
      </td>
    </tr>
    <tr>
      <td width="778" height="7" valign="top" colspan="5">
      <p align="center"><script language="javascript" src="../inc/copyright.js"></script>
      </td>
    </tr>
    <tr>
      <td width="778" height="3" valign="top" colspan="5">
      </td>
    </tr>
  </table>
  </center>
</div>
</form>
</body>

</html>
<% gzjl=htmlencode2(request("gzjl"))
if gzjl="" then Response.End
language=request("language")
lanlevel=request("lanlevel")
pthua=request("pthua")
computer=request("computer")
othertc=htmlencode2(request("othertc"))
gznum=request("gznum")
if othertc="" then othertc="无其他特长" end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from person where uname='"&uname&"'"
rs.open sql,conn,3,3
rs("gznum")=gznum
rs("gzjl")=gzjl
rs("language")=language
rs("lanlevel")=lanlevel
rs("pthua")=pthua
rs("computer")=computer
rs("othertc")=othertc
rs.update
rs.close
if modify<>"ture" then
Response.Redirect "register3.asp" 
else
Response.Redirect "modify.asp" 
end if %>







