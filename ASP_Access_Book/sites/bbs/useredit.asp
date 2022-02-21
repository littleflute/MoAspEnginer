<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<!--#include file="check.asp"-->
<%
Dim rs__MMColParam
If (Request.QueryString("id") <> "") Then 
  rs__MMColParam = Request.QueryString("id")
  '判断是否具有修改权限
  if  session("AdminUID")<>Request("id")  and session("Adminflag")<>"0" then
    response.Redirect("index.asp")
  end if
else
  response.Redirect("index.asp")
End If
%>
<%
Dim rs
Dim rs_numRows

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Source = "SELECT * FROM uers WHERE id = '" + Replace(rs__MMColParam, "'", "''") + "'"
rs.Open rs.Source,conn,1,1
rs_numRows = 0

if rs.eof then  
response.Redirect("index.asp")
response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>用户修改资料</title>
<!-- #include file="an.htm" -->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid; background-color: #e3f1d1;">
  <tr> 
    <td width="468" height="24"><a href="index.asp">论坛论坛</a>→修改资料</td>
    <td width="290">|<a href="uerlist.asp">用户列表</a> | <a href="adduer.asp">用户注册</a>| 
      <a href="find.asp">查找会员</a>| <a href="uerlist.asp?id=2">发贴排名</a>|</td>
  </tr>
</table>
  <table width="760" border="1" align="center" cellpadding="0" cellspacing="1" bordercolor="#CCCCCC">
<form action="usereditsave.asp" method="post" name="form" id="form">
    <tr > 
    <td height="21" colspan="2" background="pic/h3.gif"><div align="center"><strong><font color="#FFFFFF">修改资料</font></strong></div></td>
  </tr>
    <tr> 
      <td colspan="2" bgcolor="#FFFFFF" class="p9"><%=(rs.Fields.Item("id").Value)%>-修改资料</td>
    </tr>
    <tr> 
      <td colspan="2"> <div align="right"></div>
        <input name="id" type="hidden" id="id" value="<%=(rs.Fields.Item("id").Value)%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <p align="right">*密 码：</td>
      <td width="432"> <input name="psw" type="password" id="psw" value="<%=(rs.Fields.Item("psw").Value)%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <div align="right">请再输入一次密码：</div></td>
      <td width="432"> <input name="pswc" type="password" id="psw0" value="<%=(rs.Fields.Item("psw").Value)%>" size="20"></td>
    </tr>
    <tr> 
      <td height="18" bgcolor="#e3f1d1" width="261"> <div align="right">*性 别：</div></td>
      <td width="432"><input <%If (CStr((rs.Fields.Item("sex").Value)) = CStr("boy")) Then Response.Write("CHECKED") : Response.Write("")%> name="sex" type="radio" value="boy" checked>
        帅哥 &nbsp; <input <%If (CStr((rs.Fields.Item("sex").Value)) = CStr("girl")) Then Response.Write("CHECKED") : Response.Write("")%> type="radio" name="sex" value="girl">
        美女</td>
    </tr>
    <tr> 
      <td height="30" bgcolor="#e3f1d1"> <div align="right">个性化头像：</div></td>
      <td><select class="t2" onChange="document.images['idface'].src=options[selectedIndex].value;" size="1" name="face">
          <option value="images/01.gif">用户头像-01 </option>
          <option value="images/02.gif">用户头像-02 </option>
          <option value="images/03.gif">用户头像-03 </option>
          <option value="images/04.gif">用户头像-04 </option>
          <option value="images/05.gif">用户头像-05 </option>
          <option value="images/06.gif">用户头像-06 </option>
          <option value="images/07.gif">用户头像-07 </option>
          <option value="images/08.gif">用户头像-08 </option>
          <option value="images/09.gif">用户头像-09 </option>
          <option value="images/10.gif">用户头像-10 </option>
          <option value="images/11.gif">用户头像-11 </option>
          <option value="images/12.gif">用户头像-12 </option>
          <option value="images/13.gif">用户头像-13 </option>
          <option value="images/14.gif">用户头像-14 </option>
          <option value="images/15.gif">用户头像-15 </option>
          <option value="images/16.gif">用户头像-16 </option>
          <option value="images/17.gif">用户头像-17 </option>
          <option value="images/18.gif">用户头像-18 </option>
          <option value="images/19.gif">用户头像-19 </option>
          <option value="images/20.gif">用户头像-20</option>
        </select> <img src="<%=rs("face")%>" alt="个人形象代表" align="absmiddle" class="t2" id="idface"><font class=j1>[<a 
      href="tx.htm" target=_blank>所有头像</a>]</font></td>
    </tr>
    <tr> 
      <td height="19" bgcolor="#e3f1d1"> <div align="right">头衔：</div></td>
      <td><input name="head" type="text" id="head" value="<%=(rs.Fields.Item("head").Value)%>"></td>
    </tr>
    <tr> 
      <td height="19" bgcolor="#e3f1d1" width="261"> <div align="right">Q Q：</div></td>
      <td width="432"><input name="qq" type="text" id="qq" value="<%=(rs.Fields.Item("qq").Value)%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <div align="right">*E-mail：</div></td>
      <td width="432"><input name="email" type="text" id="email" value="<%=(rs.Fields.Item("email").Value)%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <p align="right">个人主页：</td>
      <td width="432"><input name="homepage" type="text" id="homepage" value="<%=(rs.Fields.Item("homepage").Value)%>"> 
      </td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <div align="right">院系&amp;年级：</div></td>
      <td width="432"> <input name="college" type="text" id="college" value="<%=(rs.Fields.Item("college").Value)%>"> 
      </td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <div align="right">*是否创协会员：</div></td>
      <td width="432"> <input <%If (CStr((rs.Fields.Item("huiyuan").Value)) = CStr("yes")) Then Response.Write("CHECKED") : Response.Write("")%> name="huiyuan" type="radio" value="yes" checked>
        是 &nbsp; <input <%If (CStr((rs.Fields.Item("huiyuan").Value)) = CStr("no")) Then Response.Write("CHECKED") : Response.Write("")%> type="radio" name="huiyuan" value="no">
        否</td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1">&nbsp;</td>
      <td bgcolor="#e3f1d1">&nbsp;</td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <div align="right">*姓名：</div></td>
      <td width="432"><input name="name" type="text" id="name" value="<%=(rs.Fields.Item("name").Value)%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <div align="right">部门：</div></td>
      <td width="432"> <select name="part" id="part">
          <option value="非会员" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("非会员" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>非会员</option>
          <option value="会长" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("会长" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>会长</option>
          <option value="副会长" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("副会长" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>副会长</option>
          <option value="秘书处" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("秘书处" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>秘书处</option>
          <option value="项目部" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("项目部" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>项目部</option>
          <option value="公关部" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("公关部" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>公关部</option>
          <option value="宣传部" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("宣传部" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>宣传部</option>
          <option value="OEC活动部" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("OEC活动部" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>OEC活动部</option>
          <option value="会员俱乐部" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("会员俱乐部" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>会员俱乐部</option>
          <%
While (NOT rs.EOF)
%>
          <option value="<%=(rs.Fields.Item("part").Value)%>" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If (CStr(rs.Fields.Item("part").Value) = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(rs.Fields.Item("part").Value)%></option>
          <%
  rs.MoveNext()
Wend
If (rs.CursorType > 0) Then
  rs.MoveFirst
Else
  rs.Requery
End If
%>
        </select></td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <div align="right">职务：</div></td>
      <td width="432"> <select name="class" id="class">
          <option value="非会员" <%If (Not isNull((rs.Fields.Item("class").Value))) Then If ("非会员" = CStr((rs.Fields.Item("class").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>非会员</option>
          <option value="部长" <%If (Not isNull((rs.Fields.Item("class").Value))) Then If ("部长" = CStr((rs.Fields.Item("class").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>部长</option>
          <option value="副部长" <%If (Not isNull((rs.Fields.Item("class").Value))) Then If ("副部长" = CStr((rs.Fields.Item("class").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>副部长</option>
          <option value="会员" <%If (Not isNull((rs.Fields.Item("class").Value))) Then If ("会员" = CStr((rs.Fields.Item("class").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>会员</option>
          <option value="会长" <%If (Not isNull((rs.Fields.Item("class").Value))) Then If ("会长" = CStr((rs.Fields.Item("class").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>会长</option>
          <option value="副会长" <%If (Not isNull((rs.Fields.Item("class").Value))) Then If ("副会长" = CStr((rs.Fields.Item("class").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>副会长</option>
          <%
While (NOT rs.EOF)
%>
          <option value="<%=(rs.Fields.Item("class").Value)%>" <%If (Not isNull((rs.Fields.Item("class").Value))) Then If (CStr(rs.Fields.Item("class").Value) = CStr((rs.Fields.Item("class").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(rs.Fields.Item("class").Value)%></option>
          <%
  rs.MoveNext()
Wend
If (rs.CursorType > 0) Then
  rs.MoveFirst
Else
  rs.Requery
End If
%>
        </select> </td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <div align="right">电 话：</div></td>
      <td width="432"> <input name="tel" type="text" id="tel" value="<%=(rs.Fields.Item("tel").Value)%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <div align="right">住址：</div></td>
      <td width="432"> <input name="addr" type="text" id="addr" value="<%=(rs.Fields.Item("addr").Value)%>"></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center"><br>
          <input type="submit" name="Submit" value="确 定 修 改">
        </div></td>
    </tr>
</form>  </table>
<%
rs.Close()
Set rs = Nothing
%>
<!--#include file="foot.asp"-->