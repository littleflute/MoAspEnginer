<%
Option Explicit
dim rs_lar
dim px,sql,search_txt,str
dim s_netname,s_sex,s_britherday,s_netcall,s_ip,s_renqi,s_home,s_date
dim j_netname,j_sex,j_britherday,j_netcall,j_ip,j_renqi,j_home,j_date
dim pages,records,currentpage,linenumber,line
dim p%>
<!--#include file="conn.asp"-->
<%

'�Ѷ�Session�����Ƿ�ʱ
if isempty(session("user_id")) then
   response.redirect "timeout.htm"
end if

px=request("px"):if px="" then px="j_date"

search_txt=request("search_txt")
if instr(search_txt,"%") or instr(search_txt,"'") or instr(search_txt,"<") then
	response.write "�������˷Ƿ��ַ�!"
	response.end
end if

Set rs_lar = Server.CreateObject("ADODB.Recordset")

sql="select user_id,netname,sex,netcall,britherday,ip,renqi,home,date,photo from larchives"
if search_txt<>"" then
	sql=sql & " where netname like '%" & search_txt & "%'"
	sql=sql & " or sex like '%" & search_txt & "%'"
	sql=sql & " or netcall like '%" & search_txt & "%'"
	sql=sql & " or britherday like '%" & search_txt & "%'"
	sql=sql & " or ip like '%" & search_txt & "%'"
	sql=sql & " or cstr(renqi) like '" & search_txt & "'"
	sql=sql & " or date like '%" & search_txt & "%'"
	'sql=sql & " or home like '%" & search_txt & "%'"
end if

Select case px
       case "s_netname"
             sql=sql & " Order by  time   asc"
       case "s_sex"
             sql=sql & " Order by  sex    asc"
       case "s_britherday"
             sql=sql & " Order by  britherday    asc"
       case "s_netcall"
             sql=sql & " Order by  netcall asc"
       case "s_ip"
             sql=sql & " Order by  ip      asc"
       case "s_renqi"
             sql=sql & " Order by  renqi   asc"
       case "s_home"
             'sql=sql & " Order by home asc"
       case "s_date"
             sql=sql & " Order by date     asc"
       case "s_photo"
       	     sql=sql & " Order by photo     asc"
       
       case "j_netname"
             sql=sql & " Order by  name   desc"
       case "j_sex"
             sql=sql & " Order by  sex    desc"
       case "j_britherday"
             sql=sql & " Order by  britherday    desc"
       case "j_netcall"
             sql=sql & " Order by netcall desc"
       case "j_ip"
             sql=sql & " Order by  ip     desc"
       case "j_renqi"
             sql=sql & " Order by  renqi  desc"
       case "j_home"
             'sql=sql & " Order by home    desc"
       case "j_date"
             sql=sql & " Order by time    desc"
       case "j_photo"
             sql=sql & " Order by photo    desc"
End Select
rs_lar.open sql,conn,1,1
if rs_lar.eof and rs_lar.bof then str="�˽���ϵͳ��û��һ������!"

'��ҳ����
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
<title>�����б�</title>
<meta name="generator" content="Namo WebEditor v4.0(Trial)">
<style type="text/css">A:visited {
 COLOR: #000000; TEXT-DECORATION: none
}
A:link {
 COLOR: #000000; TEXT-DECORATION: none
}
A:hover {
 COLOR: #0080c0; TEXT-DECORATION: none
}
TD {
 FONT-SIZE: 9pt; COLOR: #000000; FONT-FAMILY: "����"
}
.f1 {
 LINE-HEIGHT: 18px
}
.en {
 FONT-WEIGHT: bold; FONT-FAMILY: "Arial","Verdana"
}
.new {
 FONT-WEIGHT: bold; COLOR: #ff3300; FONT-FAMILY: "Arial"
}
.line {
 LINE-HEIGHT: 19px
}
</style>
<style type="text/css">
<!--
body {
	background-image: url(bihaibg.jpg);
}
-->
</style>
<SCRIPT language=JavaScript>
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  if (theURL != "fuckyou")
 {   window.open(theURL,winName,features);}
}
//-->
</SCRIPT></head>
<body leftmargin="5" topmargin="0">
<table width="772" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="images/top.gif" width="772" height="112" usemap="#Map" border="0"></td>
  </tr>
</table>
<table width="776" border="0" align="center" cellpadding="2" cellspacing="0">
  <tr> 
    <td width="537" align="center" style="border-bottom-style: solid; border-bottom-width: 1" height="18"> 
      <p align="left"> <font size="2"><%if search_txt<>"" then%>�����ǰ����ؼ���<font color="#ff0000"><%=search_txt%></font>����������<%else%>�������������ѵ��б�<%end if%></font> 
      </p>
    </td>
    <td width="231" align="center" style="border-bottom-style: solid; border-bottom-width: 1" height="18">&nbsp; 
    </td>
  </tr>
  <tr> 
    <td height="18" colspan="2"> <font size="2">��<font color="red"><%=currentpage%></font>ҳ&nbsp;<%if search_txt="" then%>�Ѽ������ѹ�<%else%>�ҵ�����<%end if%><font color="blue"><%=Records%></font>λ&nbsp;��<font color="blue"><%=Pages%></font>ҳ</font> 
    </td>
  </tr>
  <tr align="right"> <%if pages>1 then%> 
    <td height="18" colspan="2"> <font size="2">ת��ҳ��:[ <%for p=1 to pages%> <a <%if currentpage=p then%> style="color:red" <%end if%> href="list.asp?currentpage=<%=p%>&search_txt=<%=search_txt%>&px=<%=px%>"><%=p%></a> 
      <%if p/30=0 then%> <%end if%> <%next%> ]</font> </td>
    <%end if%> </tr>
</table>
  <!------------------------------------------------------------------------------------------------------------------------->
    
<table width="776" border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#333300" bordercolordark="#FFFFFF">
  <tr bgcolor="#336600"> 
    <td width="85" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><font color="#FFFFFF">����<a title="��������������" <%if px="s_netname" then%> style="color:red" <%end if%> href="list.asp?px=s_netname&search_txt=<%=search_txt%>">��</a><a title="��������������" <%if px="j_netname" then%> style="color:red" <%end if%> href="list.asp?px=j_netname&search_txt=<%=search_txt%>">��</a></font></td>
    <td width="59" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><font color="#FFFFFF">�Ա�<a title="���Ա���������" <%if px="s_sex" then%> style="color:red" <%end if%> href="list.asp?px=s_sex&search_txt=<%=search_txt%>">��</a><a title="���Ա�������" <%if px="j_sex" then%> style="color:red" <%end if%> href="list.asp?px=j_sex&search_txt=<%=search_txt%>">��</a></font></td>
    <td width="106" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><font color="#FFFFFF">����<a title="��������������" <%if px="s_home" then%> style="color:red" <%end if%> href="list.asp?px=s_home&search_txt=<%=search_txt%>">��</a><a title="�����ή������" <%if px="j_home" then%> style="color:red" <%end if%> href="list.asp?px=j_home&search_txt=<%=search_txt%>">��</a></font></td>
    <td width="82" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><font color="#FFFFFF">����<a title="��������������" <%if px="s_britherday" then%> style="color:red" <%end if%> href="list.asp?px=s_britherday&search_txt=<%=search_txt%>">��</a><a title="�����ս�������" <%if px="j_britherday" then%> style="color:red" <%end if%> href="list.asp?px=j_britherday&search_txt=<%=search_txt%>">��</a></font></td>
    <td width="98" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><font color="#FFFFFF">����Ѱ����<a title="�����紫����������������" <%if px="s_netcall" then%> style="color:red" <%end if%> href="list.asp?px=s_netcall&search_txt=<%=search_txt%>">��</a><a title="�����紫�������뽵������" <%if px="j_netcall" then%> style="color:red" <%end if%> href="list.asp?px=j_netcall&search_txt=<%=search_txt%>">��</a></font></td>
    <td width="98" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><font color="#FFFFFF">�Ǽ�IP<a title="���Ǽ�IP��������" <%if px="s_ip" then%> style="color:red" <%end if%> href="list.asp?px=s_ip&search_txt=<%=search_txt%>">��</a><a title="���Ǽ�IP��������" <%if px="j_ip" then%> style="color:red" <%end if%> href="list.asp?px=j_ip&search_txt=<%=search_txt%>">��</a></font></td>
    <td width="86" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><font color="#FFFFFF">�Ǽ�ʱ��<a title="���Ǽ�ʱ����������" <%if px="s_date" then%> style="color:red" <%end if%> href="list.asp?px=s_date&search_txt=<%=search_txt%>">��</a><a title="���Ǽ�ʱ�併������" <%if px="j_date" then%> style="color:red" <%end if%> href="list.asp?px=j_date&search_txt=<%=search_txt%>">��</a></font></td>
    <td width="74" align="center" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><font color="#FFFFFF">��Ƭ��<a title="���Ǽ�ʱ����������" <%if px="s_date" then%> style="color:red" <%end if%> href="list.asp?px=s_photo&search_txt=<%=search_txt%>">��</a><a title="���Ǽ�ʱ�併������" <%if px="j_date" then%> style="color:red" <%end if%> href="list.asp?px=j_photo&search_txt=<%=search_txt%>">��</a></font></td>
    <td width="68" align="center" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><font color="#FFFFFF">����<a title="��������������" <%if px="s_renqi" then%> style="color:red" <%end if%> href="list.asp?px=s_renqi&search_txt=<%=search_txt%>">��</a><a title="��������������" <%if px="j_renqi" then%> style="color:red" <%end if%> href="list.asp?px=j_renqi&search_txt=<%=search_txt%>">��</a></font></td>
  </tr>
  <%linenumber=rs_lar.pagesize%> 
  <%do while (not rs_lar.eof) and (line<linenumber)%> 
  <tr bgcolor="#FFEBE1"> 
    <td width="85" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><a href="read.asp?user_id=<%=rs_lar("user_id")%>" title="�鿴����<%=rs_lar("netname")%>����ϸ����"><%=rs_lar("netname")%></a></td>
    <td width="59" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><%=rs_lar("sex")%></td>
    <td width="106" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><%=rs_lar("home")%></td>
    <td width="82" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><%=rs_lar("britherday")%></td>
    <td width="98" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><%=rs_lar("netcall")%></td>
    <td width="98" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><%=rs_lar("ip")%></td>
    <td width="86" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><%=rs_lar("date")%></td>
    <td width="74" align="center" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><%=rs_lar("photo")%></td>
    <td width="68" align="center" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-bottom-style: solid; border-bottom-width: 1" height="20"><%=rs_lar("renqi")%></td>
  </tr>
  <%rs_lar.movenext%> 
  <%line=line+1%> 
  <%loop%> 
</table>
  <!-------------------------------------��ҳ����--------------------------------------->
  
<table width="776" border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="white">
  <tr align="center"> 
    <td bgcolor="white"> <%if currentpage>1 then%> <font size="2"><a href="list.asp?currentpage=<%=currentpage-1%>&search_txt=<%=search_txt%>&px=<%=px%>"><font size="2">[��һҳ]</font></a> 
      <%end if%> <%if currentpage<pages then%> <a href="list.asp?currentpage=<%=currentpage+1%>&search_txt=<%=search_txt%>&px=<%=px%>">[��һҳ]</a> 
      <%end if%> <%if currentpage>1 then%> <a href="list.asp?currentpage=1&search_txt=<%=search_txt%>&px=<%=px%>">[����ҳ]</a> 
      <%end if%> <%if currentpage<pages then%> <a href="list.asp?currentpage=<%=pages%>&search_txt=<%=search_txt%>&px=<%=px%>">[��ĩҳ]</a> 
      <%end if%> </font> </td>
  </tr>
</table>
  <!-------------------------------------------------------------------------------------------------------------------------><!------------------------------------------------------------------------------------------------------------------------->
<form>
  
<table width="776" height="27" border="0" align="center" cellspacing="0">
  <tr>
      <td width="100%" height="13">
      <p align="center"> �����ؼ���
        <input type="text" name="search_txt" size="8" value="<%=search_txt%>">
        <input type="submit" value="����" name="B1">
      </p>
      </td>
  </tr>
    <tr>
      
    <td width="100%" height="25"> 
      <p align="center">�ؼ���Ϊ���г��������� 
    </td>
    </tr>
</table>
        
<map name="Map">
  <area shape="rect" coords="228,95,273,110" href="default.asp">
  <area shape="rect" coords="293,94,358,109" href="your.asp">
  <area shape="rect" coords="384,95,444,111" href="list.asp">
  <area shape="rect" coords="470,95,527,112" href="register.asp">
  <area shape="rect" coords="556,92,617,111" href="sendphoto.asp">
  <area shape="rect" coords="643,93,699,111" href="pics.asp">
</map>
