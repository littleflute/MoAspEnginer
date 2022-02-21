<%
if request.cookies("adminok")="" then
  response.redirect "login.asp"
end if
%>
<!--#include file="articleconn.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>信息查询系统信息修改</title>
<script LANGUAGE="javascript">
<!--

function FrmAddLink_onsubmit() {
if (document.FrmAddLink.txttitle.value=="")
	{
	  alert("Sorry，信息名称没有输入！")
	  document.FrmAddLink.txttitle.focus()
	  return false
	 }
else if(document.FrmAddLink.txturl.value=="" || document.FrmAddLink.txturl.value.toUpperCase()=="HTTP://")
	{
	  alert("Sorry，链接地址没有输入！")
	  document.FrmAddLink.txturl.focus()
	  return false
	 }

else if(document.FrmAddLink.txtcontent.value=="")
	{
	  alert("Sorry，信息简介没有输入！")
	  document.FrmAddLink.txtcontent.focus()
	  return false
	 }
else if(document.FrmAddLink.big.value=="")
	{
	  alert("Sorry，信息大小没有输入！")
	  document.FrmAddLink.big.focus()
	  return false
	 }
else if(document.FrmAddLink.from.value=="")
	{
	  alert("Sorry，相关主页没有输入！")
	  document.FrmAddLink.from.focus()
	  return false
	 }
else if(document.FrmAddLink.fromurl.value=="")
	{
	  alert("Sorry，相关来源地址没有输入！")
	  document.FrmAddLink.fromurl.focus()
	  return false
	 }

}

//-->
</script>
<link rel="stylesheet" href="/css/style.css">
</head>

<body bgcolor="#FFFFFF">
<form align="center" method="post" name="FrmAddLink" LANGUAGE="javascript"
    onsubmit="return FrmAddLink_onsubmit()" action="saveedit.asp?id=<%=request("id")%>">
  <div align="center"><center>
      <table border="1" cellspacing="0" width="90%" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <%
dim sql
dim rs
 sql="select * from learning where articleid="&request("id")
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
                %>         
<tr bgcolor="#0099CC"> 
          <td width="100%" height="20" bgcolor="#B5D85E"> 
            <p align="center"><b><font
      color="#FFFFFF"> 信 息 修 改</font></b> 
          </td>
        </tr>
        <tr align="center"> 
          <td width="100%" height="102"> 
            <table border="0" cellspacing="1" width="93%">
              <tr> 
                <td width="15%" align="right" height="9" valign="middle"><b><font color="#0080C0">信息名称：</font></b></td>
                <td height="9" colspan="3" width="85%"> 
                  <input class=TextBorder name=txttitle size=70 maxlength="255" value="<%=rs("title")%>">
                  <font color="#FF0000"> </font></td>
              </tr>
              <tr> 
                <td width="15%" align="right" height="2" valign="middle"><b><font color="#0080C0">链接地址：</font></b></td>
                <td height="2" colspan="3" width="85%"> 
                  <input class=TextBorder name=txturl size=70 maxlength="255" value="<%=rs("url")%>">
                </td>
              </tr>
              <tr> 
                <td width="15%" align="right" height="23" valign="top"><b><font color="#0080C0">信息说明：</font></b></td>
                <td height="23" colspan="3" width="85%"> 
                  <textarea class="TextBorder" name="txtcontent" cols="70" rows="5"><%content=replace(rs("content"),"<br>",chr(13))
            content=replace(content,"&nbsp;"," ")
            response.write content%></textarea>
                  <font color="#FF0000"> </font></td>
              </tr>
              <tr> 
                <td width="15%" align="right" height="19" valign="top"><font color="#0080C0"><b> 
                  </b></font><b><font color="#0080C0"> </font><font color="#FFFFFF"><b><font color="#0080C0">信息大小：</font></b></font></b></td>
                <td height="19" colspan="3" width="85%"><b><b><font color="#0080C0"> 
                  <input class=TextBorder name=big size=10 maxlength="10" value="<%=rs("big")%>">
                  </font></b></b></td>
              </tr>
              <tr> 
                <td width="15%" align="right" height="19" valign="top"><b><font color="#0080C0">相关主页：</font></b></td>
                <td height="19" colspan="3" width="85%"><b><font color="#0080C0"> 
                  </font><font color="#0080C0"> 
                  <input class=TextBorder name=from size=20 maxlength="20" value="<%=rs("from")%>">
                  </font><font color="#0099CC"> </font></b></td>
              </tr>
              <tr> 
                <td width="15%" align="right" height="19" valign="top"><b><font color="#0080C0">相关地址：</font></b></td>
                <td height="19" colspan="3" width="85%"><b><font color="#0080C0"> 
                  <input class=TextBorder name=fromurl size=20 maxlength="255" value="<%=rs("fromurl")%>">
                  </font></b></td>
              </tr>
              <tr> 
                <td colspan="4" align="right" height="19" valign="top"> 
                  <div align="center"></div>
                  <div align="center"><font color="#0099CC">ID号：<%=rs("articleid")%> 
                    类型：<%=rs("type")%> 评价：<font color="#FF0000"><%=rs("vote")%> 
                    </font>创建日期：<%=rs("dateandtime")%> 访问量：<%=rs("hits")%></font></div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </center></div><div align="center"><center><p><input type="submit" value=" 修 改 "
  name="cmdok" class="buttonface">&nbsp; <input type="reset" value=" 复 原 "
  name="cmdcancel" class="buttonface"></p>
  </center></div>
</form>
</body>
</html>
<% 
   
 rs.close 
set rs=nothing 
  conn.close  
  set conn=nothing  
%>
