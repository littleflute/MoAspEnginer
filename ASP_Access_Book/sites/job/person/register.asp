<% Response.Buffer=True %>
<!--#include file="../inc/person.inc"-->
<% uname=session("puid") 
   modify=request("modify")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from person where uname='"&uname&"'"
rs.open sql,conn,1,1 %>
<% if modify<>"ture" and rs("job") <> "" then
   response.write"<SCRIPT language=JavaScript>alert('你已经登录个人简历，请不要重复登录！');"
   response.write"javascript:history.go(-1)</SCRIPT>"
   end if %>

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
<%end if%>
</head>
<SCRIPT language=JavaScript src="../inc/validate.js"></SCRIPT>
<SCRIPT language=JavaScript src="../inc/vreg1.js"></SCRIPT>
<% if modify<>"ture" then %>
<FORM name=register action=register.asp method=post>
<%else%>
<FORM name="register" action="register.asp?modify=ture" method="post">
<%end if%>
<body topmargin="0" leftmargin="0">
<!--#include file="../inc/top2.htm"-->
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="317">
    <tr>
        <td width="778" height="18" valign="top" colspan="5" bgcolor="#53BEB0"> 
          　 </td>
    </tr>
    <tr>
        <td width="139" valign="top" bgcolor="#53BEB0"> 　 
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
          <p align="center"> <br>
        <br>
      </td>
      <td width="36" height="243" valign="top"><img border="0" src="../images/selfk.GIF"></td>
      <td width="456" height="243" valign="top">
  </center>
    
      　
    
        <table border="1" cellpadding="0" cellspacing="0" width="421" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
          <tr>
         <td width="417" height="18" valign="bottom" bgcolor="#C6CEDE" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
              <p align="center">&nbsp;=== 个人基本资料 ===</td>                            
          </tr>
          <tr>
            <td width="417" height="20" valign="bottom">
              <p align="left">&nbsp;用户名:<FONT COLOR="#FF0000"><%=uname%></FONT> &nbsp;           
            </td>
          </tr>
          <tr>
            
			<td width="417" height="20" valign="bottom" bgcolor="#F3F3F3">
              &nbsp;真实姓名:<%if modify<>"ture" then%><input type="text" name="iname" valign="bottom" size="12" maxLength=4 value="<%=rs("iname")%>">
                <%else%>
                <%=rs("iname")%> 
                <%end if%>
                (<font color="#FF0000">*</font>必填)</td>
          </tr>
          <tr>
              <td width="417" height="28"> 
                <p align="left">&nbsp;性别:
                  <input type="radio"<%if modify<>"true" or rs("sex") ="男" then Response.Write "checked"%> value="男" name="sex">
                  男 
                  <input type="radio" <%if rs("sex") ="女" then Response.Write "checked"%> name="sex" value="女">
                  女 (<font color="#FF0000">*</font>必填) </p> 
            </td>
          </tr>
          <tr>
              <td width="417" height="28" bgcolor="#F3F3F3"> 
                <p align="left">&nbsp;出生年月:<INPUT maxLength=10 size=10 name=bday maxLength=10 value="<%=rs("bday")%>"> 
                  (<font color="#FF0000">*</font>必填，例如2000-1-1) </p> 
            </td>
          </tr>
          <tr>
            <td width="417" height="28">
              &nbsp;身份证号码:<input type="text" name="code" size="18" maxLength=18 value="<%=rs("code")%>">
                (<font color="#FF0000">*</font>必填) </td>
          </tr>
          <tr>
              <td width="417" height="28" bgcolor="#F3F3F3"> 
                <% if modify<>"ture" then %>
                <p align="left">&nbsp;民族:
                  <input type="text" name="mzhu" size="12" maxLength=12 value="汉族">
                  (<font color="#FF0000">*</font>必填) 
                  <%else%>
                <p align="left">&nbsp;民族:
                  <input type="text" name="mzhu" size="12" maxLength=12 value="<%=rs("mzhu")%>">
                  (<font color="#FF0000">*</font>必填) 
                  <%end if%>
                </p> 
            </td>
          </tr>
          <tr>
              <td width="417" height="28"> 
                <p align="left">&nbsp;婚姻状况:
                  <SELECT size=1 name=marry>
                    <OPTION value=未婚 <%if rs("marry") ="未婚" then Response.Write "selected"%>>未婚 
                    <OPTION value=已婚 <%if rs("marry") ="已婚" then Response.Write "selected"%>>已婚 
                    <OPTION value=离异 <%if rs("marry") ="离异" then Response.Write "selected"%>>离异</OPTION>
                  </SELECT>
                  (<font color="#FF0000">*</font>必填)</p>
            </td>
          </tr>
          <tr>
              <td width="417" height="28" bgcolor="#F3F3F3"> 
                <p align="left">&nbsp;户籍所在地:
                  <SELECT size=1 name=hka>
                    <OPTION>请在以下列表中选择</OPTION>
                    <OPTION value=北京市 <%if rs("hka") ="北京市" then Response.Write "selected"%>>北京市 
                    <OPTION value=天津市 <%if rs("hka") ="天津市" then Response.Write "selected"%>>天津市 
                    <OPTION value=上海市 <%if rs("hka") ="上海市" then Response.Write "selected"%>>上海市 
                    <OPTION value=重庆市 <%if rs("hka") ="重庆市" then Response.Write "selected"%>>重庆市 
                    <OPTION value=安徽省 <%if rs("hka") ="安徽省" then Response.Write "selected"%>>安徽省 
                    <OPTION value=福建省 <%if rs("hka") ="福建省" then Response.Write "selected"%>>福建省 
                    <OPTION value=甘肃省 <%if rs("hka") ="甘肃省" then Response.Write "selected"%>>甘肃省 
                    <OPTION value=广东省 <%if rs("hka") ="广东省" then Response.Write "selected"%>>广东省 
                    <OPTION value=四川省 <%if rs("hka") ="四川省" then Response.Write "selected"%>>四川省 
                    <OPTION value=贵州省 <%if rs("hka") ="贵州省" then Response.Write "selected"%>>贵州省 
                    <OPTION value=湖北省 <%if rs("hka") ="湖北省" then Response.Write "selected"%>>湖北省 
                    <OPTION value=湖南省 <%if rs("hka") ="湖南省" then Response.Write "selected"%>>湖南省 
                    <OPTION value=吉林省 <%if rs("hka") ="吉林省" then Response.Write "selected"%>>吉林省 
                    <OPTION value=江苏省 <%if rs("hka") ="江苏省" then Response.Write "selected"%>>江苏省 
                    <OPTION value=江西省 <%if rs("hka") ="江西省" then Response.Write "selected"%>>江西省 
                    <OPTION value=辽宁省 <%if rs("hka") ="辽宁省" then Response.Write "selected"%>>辽宁省 
                    <OPTION value=青海省 <%if rs("hka") ="青海省" then Response.Write "selected"%>>青海省 
                    <OPTION value=山东省 <%if rs("hka") ="山东省" then Response.Write "selected"%>>山东省 
                    <OPTION value=山西省 <%if rs("hka") ="山西省" then Response.Write "selected"%>>山西省 
                    <OPTION value=陕西省 <%if rs("hka") ="陕西省" then Response.Write "selected"%>>陕西省 
                    <OPTION value=云南省 <%if rs("hka") ="云南省" then Response.Write "selected"%>>云南省 
                    <OPTION value=浙江省 <%if rs("hka") ="浙江省" then Response.Write "selected"%>>浙江省 
                    <OPTION value=黑龙江省 <%if rs("hka") ="黑龙江省" then Response.Write "selected"%>>黑龙江省 
                    
                  </SELECT>
                  (<font color="#FF0000">*</font>必填)</p>
            </td>
          </tr>
          <tr>
              <td width="417" height="28"> 
                <p align="left"> &nbsp;您最高的教育程度:
                  <SELECT size=1 
                  name=edu>
                    <OPTION>请在以下列表中选择</OPTION>
                    <OPTION value=初中 <%if rs("edu") ="初中" then Response.Write "selected"%>>初中 
                    <OPTION value=高中 <%if rs("edu") ="高中" then Response.Write "selected"%>>高中 
                    <OPTION value=中技 <%if rs("edu") ="中技" then Response.Write "selected"%>>中技 
                    <OPTION value=中专 <%if rs("edu") ="中专" then Response.Write "selected"%>>中专 
                    <OPTION value=大专 <%if rs("edu") ="大专" then Response.Write "selected"%>>大专 
                    <OPTION value=本科 <%if rs("edu") ="本科" then Response.Write "selected"%>>本科 
                    <OPTION value=硕士 <%if rs("edu") ="硕士" then Response.Write "selected"%>>硕士 
                    <OPTION value=博士 <%if rs("edu") ="博士" then Response.Write "selected"%>>博士</OPTION>
                  </SELECT>
                  (<font color="#FF0000">*</font>必填)</p> 
            </td>
          </tr>
          <tr>
            <td width="417" height="28" bgcolor="#F3F3F3">
              <p align="left"> 
              &nbsp;专 业:<INPUT maxLength=60 size=30                            
                  name=zye maxLength=30 value="<%=rs("zye")%>">(<font color="#FF0000">*</font>必填)</p> 
            </td>
          </tr>
          <tr>
            <td width="417" height="28">
              <p align="left"> 
              &nbsp;毕业院校:<INPUT maxLength=60 size=40 
                  name=school maxLength=40 value="<%=rs("school")%>">(<font color="#FF0000">*</font>必填)</p> 
            </td>
          </tr>
          <tr>
              <td width="417" height="28" bgcolor="#F3F3F3"> 
                <p align="left"> &nbsp;政治面貌:
                  <SELECT size=1 name=zzmm>
                    <OPTION>请在以下列表中选择</OPTION>
                    <OPTION value=党员 <%if rs("zzmm") ="党员" then Response.Write "selected"%>>党员 
                    <OPTION value=团员 <%if rs("zzmm") ="团员" then Response.Write "selected"%>>团员 
                    <OPTION value=群众 <%if rs("zzmm") ="群众" then Response.Write "selected"%>>群众 
                    <OPTION value=民主党派 <%if rs("zzmm") ="民主党派" then Response.Write "selected"%>>民主党派 
                    <OPTION value=其它 <%if rs("zzmm") ="其它" then Response.Write "selected"%>>其它</OPTION>
                  </SELECT>
                  (<font color="#FF0000">*</font>必填)</p> 
            </td>
          </tr>
          <tr>
              <td width="417" height="28"> 
                <p align="left"> &nbsp;现有职称:
                  <SELECT size=1 name=zchen>
                    <OPTION>请在以下列表中选择</OPTION>
                    <OPTION value=高级 <%if rs("zchen") ="高级" then Response.Write "selected"%>>高级 
                    <OPTION value=中级 <%if rs("zchen") ="中级" then Response.Write "selected"%>>中级 
                    <OPTION value=初级 <%if rs("zchen") ="初级" then Response.Write "selected"%>>初级 
                    <OPTION value=暂无 <%if rs("zchen") ="暂无" then Response.Write "selected"%>>暂无 
                    <OPTION value=其它 <%if rs("zchen") ="其它" then Response.Write "selected"%>>其它</OPTION>
                  </SELECT>
                  (<font color="#FF0000">*</font>必填)</p> 
            </td>
          </tr>
          <tr>
            <td valign="top" width="417">
              
              <p align="center"> 
              <br>
              <% if modify<>"ture" then %>
			  <input type="button" value="下一步" onClick="check()"></p>  
              <%else%>
              <input type="button" value="更 新" onClick="checkmod()">  
             <br>
             <%end if%>      
            <br>
            </td>
          </tr>
        </table>
      </td>
  <center>
      <td width="1" height="243" valign="top" bgcolor="#00006A"></td>
      <td width="138" height="243" valign="top" bgcolor="#F3F3F3">　</td>
    </tr>
    <tr>
        <td width="778" height="1" valign="top" colspan="5" bgcolor="#53BEB0"> 
          <p align="center"> </td>
    </tr>
    <tr>
      <td width="778" height="4" valign="top" colspan="5">
      <p align="center">
      </td>
    </tr>
    <tr>
      <td width="778" height="4" valign="top" colspan="5">
      <p align="center"><script language="javascript" src="../inc/copyright.js"></script>
      </td>
    </tr>
    <tr>
      <td width="778" height="4" valign="top" colspan="5">
      </td>
    </tr>
  </table>
  </center>
</div>
</form>
</body>

</html>
<% bday=request("bday") 
if bday="" then Response.End
sex=request("sex")
high=request("high")
iname=request("iname")
code=request("code")
mzhu=request("mzhu")
marry=request("marry")
hka=request("hka")
zye=request("zye")
edu=request("edu")
school=request("school")
zzmm=request("zzmm")
zchen=request("zchen")
if iname="" then iname=rs("iname") end if
rs.close
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from person where uname='"&uname&"'"
rs.open sql,conn,3,3
rs("iname")=iname
rs("code")=code
rs("mzhu")=mzhu
rs("marry")=marry
rs("zzmm")=zzmm
rs("zchen")=zchen
rs("bday")=bday
rs("sex")=sex
rs("hka")=hka
rs("zye")=zye
rs("edu")=edu
rs("school")=school
rs.update
rs.close
if modify<>"ture" then
Response.Redirect "register2.asp" 
else
Response.Redirect "modify.asp" 
end if %>