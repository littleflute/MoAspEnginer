<!--#include file="conn.asp"-->
<%if session("admin_pass")<>"ok" then
   response.redirect "adminlogin.asp"
   response.end
end if
px=request("px")
if px="" then px="j_date"
search_txt=request("search_txt")
Set rs_lar = Server.CreateObject("ADODB.Recordset")

sql="select user_id,netname,sex,netcall,britherday,ip,renqi,home,date from larchives"
if search_txt<>"" then
	sql=sql & " where netname like '%" & search_txt & "%'"
	sql=sql & " or sex like '%" & search_txt & "%'"
	sql=sql & " or netcall like '%" & search_txt & "%'"
	sql=sql & " or britherday like '%" & search_txt & "%'"
	sql=sql & " or ip like '%" & search_txt & "%'"
	sql=sql & " or renqi like '%" & search_txt & "%'"
	sql=sql & " or date like '" & search_txt & "'"
	sql=sql & " or home like '%" & search_txt & "%'"
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
             sql=sql & " Order by home     asc"
       case "s_date"
             sql=sql & " Order by date     asc"

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
             sql=sql & " Order by home    desc"
       case "j_date"
             sql=sql & " Order by time    desc"
End Select

rs_lar.open sql,conn,1,1
if rs_lar.eof and rs_lar.bof then str="�˽���ϵͳ��û��һ������!"
'��ҳ����
if str="" then
	rs_lar.PageSize=30 'ÿҳ����
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
-->
</style>
<style type="text/css">
<!--
body {
	background-image: url(bihaibg.jpg);
}
-->
</style>
<meta name="generator" content="Namo WebEditor v4.0(Trial)">
<link rel="stylesheet" href="ysb.css">
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>
<body topmargin="0" leftmargin="5" onLoad="MM_openBrWindow(','','toolbar=yes,location=yes,status=yes,menubar=yes,scrollbars=yes,resizable=yes,width=500,height=450')">
<table width="772" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="images/top.gif" width="772" height="112" usemap="#Map" border="0"></td>
  </tr>
</table>
<table width="772" border="0" align="center">
  <tr>
    <td width="578" align="center" style="border-bottom-style: solid; border-bottom-width: 1">
      <table border="0" width="100%" cellspacing="0" cellpadding="0" bgcolor="#FFCCCC">
        <tr> 
          <td style="border-bottom-style: solid" width="88" height="19" align="center"><a href="default.asp">������ҳ</a></td>
          <td style="border-bottom-style: solid" width="41" height="19" align="center">&nbsp;</td>
          <td style="border-bottom-style: solid" width="150" height="19" align="center"><a href="exitadmin.asp" class="x"><b>�˳�����</b></a></td>
          <td style="border-bottom-style: solid" width="315" height="19">&nbsp;</td>
          <td style="border-bottom-style: solid" width="155" height="19">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td width="578" align="center" style="border-bottom-style: solid; border-bottom-width: 1"> 
      <p align="left"> <font size="2"><i><%if search_txt<>"" then%>�����ǰ����ؼ���<font color="#ff0000"><%=search_txt%></font>����������<%else%>�������������ѵ��б�</i><%end if%></font> 
      </p>
    </td>
  </tr>
  <tr> 
    <td width="578"> <font size="2">��<font color="red"><%=currentpage%></font>ҳ&nbsp;<%if search_txt="" then%>�Ѽ������ѹ�<%else%>�ҵ�����<%end if%><font color="blue"><%=Records%></font>λ&nbsp;��<font color="blue"><%=Pages%></font>ҳ</font> 
    </td>
  </tr>
  <tr> <%if pages>1 then%> 
    <td width="578"> <font size="2">ת��ҳ��:[ <%for p=1 to pages%> <a <%if currentpage=p then%> style="color:red" <%end if%> href="admin.asp?currentpage=<%=p%>&search_txt=<%=search_txt%>&px=<%=px%>"><%=p%></a> 
      <%if p/30=0 then%> <%end if%> <%next%> ]</font> </td>
    <%end if%> </tr>
</table>  
  <form name="form1">  
    
  <table width="772" border="0" align="center" cellspacing="0">
    <tr>  
        <td width="87" align="center" bgcolor="#000000" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><font color="#FFFFFF">����<a title="��������������" <%if px="s_netname" then%> style="color:red" <%end if%> href="admin.asp?px=s_netname&search_txt=<%=search_txt%>">��</a><a title="��������������" <%if px="j_netname" then%> style="color:red" <%end if%> href="admin.asp?px=j_netname&search_txt=<%=search_txt%>">��</a></font></td>  
        <td width="60" align="center" bgcolor="#000000" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><font color="#FFFFFF">�Ա�<a title="���Ա���������" <%if px="s_sex" then%> style="color:red" <%end if%> href="admin.asp?px=s_sex&search_txt=<%=search_txt%>">��</a><a title="���Ա�������" <%if px="j_sex" then%> style="color:red" <%end if%> href="admin.asp?px=j_sex&search_txt=<%=search_txt%>">��</a></font></td>  
        <td width="116" align="center" bgcolor="#000000" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><font color="#FFFFFF">����<a title="��������������" <%if px="s_home" then%> style="color:red" <%end if%> href="admin.asp?px=s_home&search_txt=<%=search_txt%>">��</a><a title="�����ή������" <%if px="j_home" then%> style="color:red" <%end if%> href="admin.asp?px=j_home&search_txt=<%=search_txt%>">��</a></font></td>  
        <td width="80" align="center" bgcolor="#000000" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><font color="#FFFFFF">����<a title="��������������" <%if px="s_britherday" then%> style="color:red" <%end if%> href="admin.asp?px=s_britherday&search_txt=<%=search_txt%>">��</a><a title="�����ս�������" <%if px="j_britherday" then%> style="color:red" <%end if%> href="admin.asp?px=j_britherday&search_txt=<%=search_txt%>">��</a></font></td>  
        <td width="99" align="center" bgcolor="#000000" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><font color="#FFFFFF">����Ѱ����<a title="�����紫����������������" <%if px="s_netcall" then%> style="color:red" <%end if%> href="admin.asp?px=s_netcall&search_txt=<%=search_txt%>">��</a><a title="�����紫�������뽵������" <%if px="j_netcall" then%> style="color:red" <%end if%> href="admin.asp?px=j_netcall&search_txt=<%=search_txt%>">��</a></font></td>  
        <td width="110" align="center" bgcolor="#000000" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><font color="#FFFFFF">�Ǽ�IP<a title="���Ǽ�IP��������" <%if px="s_ip" then%> style="color:red" <%end if%> href="admin.asp?px=s_ip&search_txt=<%=search_txt%>">��</a><a title="���Ǽ�IP��������" <%if px="j_ip" then%> style="color:red" <%end if%> href="admin.asp?px=j_ip&search_txt=<%=search_txt%>">��</a></font></td>  
        <td width="81" align="center" bgcolor="#000000" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><font color="#FFFFFF">�Ǽ�ʱ��<a title="���Ǽ�ʱ����������" <%if px="s_date" then%> style="color:red" <%end if%> href="admin.asp?px=s_date&search_txt=<%=search_txt%>">��</a><a title="���Ǽ�ʱ�併������" <%if px="j_date" then%> style="color:red" <%end if%> href="admin.asp?px=j_date&search_txt=<%=search_txt%>">��</a></font></td>  
        <td width="55" align="center" bgcolor="#000000" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><font color="#FFFFFF">����<a title="��������������" <%if px="s_renqi" then%> style="color:red" <%end if%> href="admin.asp?px=s_renqi&search_txt=<%=search_txt%>">��</a><a title="��������������" <%if px="j_renqi" then%> style="color:red" <%end if%> href="admin.asp?px=j_renqi&search_txt=<%=search_txt%>">��</a></font></td>  
        <td width="44" align="center" bgcolor="#000000" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><font color="#FFFFFF">ɾ��</font></td>  
    </tr>  
      <%linenumber=rs_lar.pagesize%>  
      <%do while (not rs_lar.eof) and (line<linenumber)%>  
      <tr>  
        <td width="87" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000" bgcolor="#EEEEEE"><a href="adminread.asp?user_id=<%=rs_lar("user_id")%>" title="�鿴����<%=rs_lar("netname")%>����ϸ����"><%=rs_lar("netname")%></a></td>  
        <td width="60" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000" bgcolor="#EEEEEE"><%=rs_lar("sex")%></td>  
        <td width="116" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000" bgcolor="#EEEEEE"><%=rs_lar("home")%></td>  
        <td width="80" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000" bgcolor="#EEEEEE"><%=rs_lar("britherday")%></td>  
        <td width="99" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000" bgcolor="#EEEEEE"><%=rs_lar("netcall")%></td>  
        <td width="110" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000" bgcolor="#EEEEEE"><%=rs_lar("ip")%></td>  
        <td width="81" align="center" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000" bgcolor="#EEEEEE"><%=rs_lar("date")%></td>  
        <td width="55" align="center" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000" bgcolor="#EEEEEE"><%=rs_lar("renqi")%></td>  
        <td width="44" align="center" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000" bgcolor="#EEEEEE"><a href="deluser.asp?user_id=<%=rs_lar("user_id")%>">[ɾ��]</a></td>  
      </tr>  
      <%rs_lar.movenext%>  
      <%line=line+1%>  
      <%loop%>  
      <%rs_lar.close%>  
    </table>  
</form>  
  
<table width="772" align="center">
  <tr>  
      <td>  
        
      <%if currentpage>1 then%>  
      <font size="2"><a href="admin.asp?currentpage=<%=currentpage-1%>&search_txt=<%=search_txt%>&px=<%=px%>"><font size="2">[��һҳ]</font></a>  
      <%end if%>  
      <%if currentpage<pages then%>  
      <a href="admin.asp?currentpage=<%=currentpage+1%>&search_txt=<%=search_txt%>&px=<%=px%>">[��һҳ]</a>  
      <%end if%>  
      <%if currentpage>1 then%>  
      <a href="admin.asp?currentpage=1&search_txt=<%=search_txt%>&px=<%=px%>">[����ҳ]</a>  
      <%end if%>  
      <%if currentpage<pages then%>  
      <a href="admin.asp?currentpage=<%=pages%>&search_txt=<%=search_txt%>&px=<%=px%>">[��ĩҳ]</a>  
      <%end if%>  
      </font>  
      </td>  
  </tr>  
</table>  
  <form>  
  <table width="772" height="27" border="0" align="center" cellspacing="0">
    <tr>  
      <td width="100%" height="13">  
      <p align="center">  
      �ؼ���<input type="text" name="search_txt" size="8"><input type="submit" value="����" name="B1">
        </p>  
      </td>  
   </tr>  
    <tr>  
      <td width="100%" height="12">  
      <p align="center">�ؼ���Ϊ���г���������  
      </td>  
    </tr>  
  </table>  
</form>  
<map name="Map">
  <area shape="rect" coords="231,95,276,110" href="default.asp">
  <area shape="rect" coords="299,95,357,112" href="your.asp">
  <area shape="rect" coords="384,93,445,111" href="list.asp">
  <area shape="rect" coords="470,95,532,112" href="register.asp">
  <area shape="rect" coords="554,95,616,111" href="sendphoto.asp">
  <area shape="rect" coords="642,94,700,112" href="pics.asp">
</map>
</body>  
</html>