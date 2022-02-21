<%Option Explicit%>
<!--#include file="conn.asp"-->
<%
dim rs_lar
dim c_age,c_name,c_sex,c_netname,c_home,c_education,c_job,c_netcall,c_tel,c_book,c_music,c_sport,c_interest,c_character,c_photo
dim age1,age2,name,sex,netname,home,education,job,netcall,tel,sport,book,music,interest,character,photo
dim sql,str,hrefdate
dim pages,records,currentpage,linenumber,line,p

c_age        =request("c_age")
c_name       =request("c_name")
c_sex        =request("c_sex")
c_netname    =request("c_netname")
c_home       =request("c_home")
c_education  =request("c_education")
c_job        =request("c_job")
c_netcall    =request("c_netcall")
c_tel        =request("c_tel")
c_sport      =request("c_sport")
c_book       =request("c_book")
c_music      =request("c_music")
c_sport      =request("c_sport")
c_interest   =request("c_interest")
c_character  =request("c_character")
c_photo      =request("c_photo")

age1         =request("age1")
age2         =request("age2")
name         =request("name")
sex          =request("sex")
netname      =request("netname")
home         =request("home")
education    =request("education")
job          =request("job")
netcall      =request("netcall")
tel          =request("tel")
sport        =request("sport")
book         =request("book")
music        =request("music")
sport        =request("sport")
interest     =request("interest")
character    =request("character")
photo        =request("photo")

Set rs_lar = Server.CreateObject("ADODB.Recordset")

hrefdate=""
sql="select * from larchives where lar_id like '%'"
if c_age="ON" then
   sql=sql & " and age>" & clng(age1) & " and age<" & clng(age2)
   hrefdate=hrefdate & "&c_age=" & "ON" & "&age1=" & age1 & "age2=" & age2
end if
if c_name="ON" then
   sql=sql & " and name like '" & name & "'"
   hrefdate=hrefdate & "&c_age=" & "ON" & "&age1=" & age1 & "age2=" & age2
end if
if c_netname="ON" then
   sql=sql & " and netname like '" & netname & "'"
   hrefdate=hrefdate & "&c_netname=" & "ON" & "&netname=" & netname
end if
if c_sex="ON" then
   sql=sql & " and sex like '" & sex & "'"
   hrefdate=hrefdate & "&c_sex=" & "ON" & "&sex=" & sex
end if
if c_home="ON" then
   sql=sql & " and home like '%" & home & "%'"
   hrefdate=hrefdate & "&c_home=" & "ON" & "&home=" & home
end if
if c_education="ON" then
   sql=sql & " and education like '" & education & "'"
   hrefdate=hrefdate & "&c_education=" & "ON" & "&education=" & education
end if
if c_job="ON" then
   sql=sql & " and job like '" & job & "'"
   hrefdate=hrefdate & "&c_job=" & "ON" & "&job=" & job
end if
if c_netcall="ON" then
   sql=sql & " and netcall like '%" & netcall & "%'"
   hrefdate=hrefdate & "&c_netcall=" & "ON" & "&netcall=" & netcall
end if
if c_tel="ON" then
   sql=sql & " and tel like '" & tel & "'"
   hrefdate=hrefdate & "&c_tel=" & "ON" & "&tel=" & tel
end if
if c_sport="ON" then
   sql=sql & " and sport like '%" & sport & "%'"
   hrefdate=hrefdate & "&c_sport=" & "ON" & "&sport=" & sport
end if
if c_book="ON" then
   sql=sql & " and book like '%" & book & "%'"
   hrefdate=hrefdate & "&c_book=" & "ON" & "&book=" & book
end if
if c_music="ON" then
   sql=sql & " and music like '%" & music & "%'"
   hrefdate=hrefdate & "&c_music=" & "ON" & "&music=" & music
end if
if c_interest="ON" then
   sql=sql & " and interest like '%" & interest & "%'"
   hrefdate=hrefdate & "&c_interest=" & "ON" & "&interest=" & interest
end if
if c_character="ON" then
   sql=sql & " and character like '%" & character & "%'"
   hrefdate=hrefdate & "&c_character=" & "ON" & "&character=" & character
end if

if c_photo="ON" then
   if photo="有" then sql=sql & " and photo>0" else sql=sql & " and photo=0"
   hrefdate=hrefdate & "&c_photo=" & "ON" & "&photo=" & photo
end if
rs_lar.open sql,conn,1,1
if rs_lar.eof and rs_lar.bof then str="此交友系统还没有一个网友!"
'分页设置
if str="" then
	rs_lar.PageSize=25
	pages=rs_lar.pagecount
	records=rs_lar.recordcount
	currentpage=request("currentpage")
	if currentpage="" or currentpage<1 then currentpage=1
	currentpage=cint(currentpage)
	if currentpage>pages then currentpage=pages
	rs_lar.absolutepage=currentpage
else
    currentpage=1
	records=0
	pages=1
end if
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Namo WebEditor v4.0(Trial)">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>搜索</title>
<style>
<!--
.x:link      { color: white; text-decoration: none }
.x:visited   { color: white; text-decoration: none }
.x:active    { color: red; text-decoration: none }
.x:hover     { color: red;text-decoration:nome}
tr           { font-size: 10pt }
body         { font-size: 10pt }
a:link       { color: blue; text-decoration: none }
a:visited    { color: blue; text-decoration: none }
a:active     { color: red; text-decoration: none }
a:hover      { color: red; text-decoration: none }
-->
</style>
<style type="text/css">
<!--
body {
	background-image: url(bihaibg.jpg);
}
-->
</style>
</head>

<body leftmargin="5" topmargin="0">
<table width="772" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="images/top.gif" width="772" height="112" usemap="#Map" border="0"></td>
  </tr>
</table>
<table width="776" border="0" align="center" cellpadding="2" cellspacing="1">
  <tr> 
    <td align="center" style="border-bottom-style: solid; border-bottom-width: 1"> 
      <p align="left"> &nbsp;以下是符合您搜索条件的网友 </p>
    </td>
  </tr>
  <tr> 
    <td height="8"> 第<font color="#CC0000"><%=currentpage%></font>页&nbsp;找到网友<font color="#CC0033"><b><%=Records%></b></font>位&nbsp;共<font color="blue"><b><font color="#FF0033"><%=Pages%></font></b></font>页 
      转到页码:[ <%for p=1 to pages%> <a <%if currentpage=p then%> style="color:red" <%end if%> href="searchout.asp?currentpage=<%=p%><%=hrefdate%>"><%=p%></a> 
      <%if p/30=0 then%> <%end if%> <%next%> ] </td>
  </tr>
</table> 
  <!-------------------------------------------------------------------------------------------------------------------------> 
  <!-------------------------------------------------------------------------------------------------------------------------> 
    
<table width="776" border="1" align="center" cellpadding="1" cellspacing="0" bordercolorlight="#CC0000" bordercolordark="#FFFFFF">
  <tr bgcolor="#336600"> 
    <td width="87" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1"><font color="#FFFFFF">网名</font></td>
    <td width="60" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1"><font color="#FFFFFF">性别</font></td>
    <td width="127" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1"><font color="#FFFFFF">来自</font></td>
    <td width="78" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1"><font color="#FFFFFF">生日</font></td>
    <td width="110" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1"><font color="#FFFFFF">网络寻呼机</font></td>
    <td width="90" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1"><font color="#FFFFFF">登记IP</font></td>
    <td width="81" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1"><font color="#FFFFFF">登记时间</font></td>
    <td width="81" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1"><font color="#FFFFFF">人气</font></td>
    <td width="69" align="center" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-bottom-style: solid; border-bottom-width: 1"><font color="#FFFFFF">照片数</font></td>
  </tr>
  <%linenumber=rs_lar.pagesize%> <%do while (not rs_lar.eof) and (line<linenumber)%> 
  <tr bgcolor="#FFEBE1"> 
    <td width="87" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1"><a href="read.asp?user_id=<%=rs_lar("user_id")%>" title="查看网友<%=rs_lar("netname")%>的祥细资料"><%=rs_lar("netname")%></a></td>
    <td width="60" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1"><%=rs_lar("sex")%></td>
    <td width="127" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1"><%=rs_lar("home")%></td>
    <td width="78" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1"><%=rs_lar("britherday")%></td>
    <td width="110" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1"><%=rs_lar("netcall")%></td>
    <td width="90" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1"><%=rs_lar("ip")%></td>
    <td width="81" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1"><%=rs_lar("date")%></td>
    <td width="81" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1"><%=rs_lar("renqi")%></td>
    <td width="69" align="center" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-bottom-style: solid; border-bottom-width: 1"><%=rs_lar("photo")%></td>
  </tr>
  <%rs_lar.movenext%> <%line=line+1%> <%loop%> <%rs_lar.close%> 
</table> 
  <!-------------------------------------分页导航---------------------------------------> 
  
<table width="776" align="center">
  <tr> 
      <td> 
       
      <%if currentpage>1 then%> 
      <font size="2"><a href="searchout.asp?currentpage=<%=currentpage-1%>"><font size="2">[上一页]</font></a> 
      <%end if%> 
      <%if currentpage<pages then%> 
      <a href="searchout.asp?currentpage=<%=currentpage+1%>">[下一页]</a> 
      <%end if%> 
      <%if currentpage>1 then%> 
      <a href="searchout.asp?currentpage=1">[最首页]</a> 
      <%end if%> 
      <%if currentpage<pages then%> 
      <a href="searchout.asp?currentpage=<%=pages%>">[最末页]</a> 
      <%end if%> 
      </font> 
      </td> 
  </tr> 
</table> 
  <!------------------------------------------------------------------------------------------------------------------------->  
<div align="left"></div>
<map name="Map">
  <area shape="rect" coords="228,93,273,114" href="default.asp">
  <area shape="rect" coords="294,96,356,111" href="your.asp">
  <area shape="rect" coords="383,94,443,113" href="list.asp">
  <area shape="rect" coords="467,95,529,111" href="register.asp">
  <area shape="rect" coords="551,93,616,112" href="sendphoto.asp">
  <area shape="rect" coords="642,92,703,113" href="pics.asp">
</map>
</body>

</html>
