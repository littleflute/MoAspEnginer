<!--#include file="conn.asp"-->
<!--#include file="isteacher.asp"-->
<!--#include file="fenlei.asp"-->
<%
id = request("id")

if id = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('�벻Ҫ����');top.window.location.href='teachermain.asp';</script>"
 response.end
end if

sql = "select * from main,teacher where main.idofteacher=teacher.teacherid and main.mainid="&id
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('�벻Ҫ����');top.window.location.href='teachermain.asp';</script>"
 response.end
else
 if rs("teacherid") <> int(session("teacherid")) and session("admin") <> "admin" then
  rs.close
  set rs = nothing
  conn.close
  set conn = nothing
  response.write "<script>alert('������ϲ����㷢���ģ������ʲô��');top.window.location.href='teachermain.asp';</script>"
  response.end
 end if
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE>TD {
	FONT-SIZE: 9pt; LINE-HEIGHT: 140%
}
BODY {
	FONT-SIZE: 9pt; LINE-HEIGHT: 140%
}
A:link {
	COLOR: #0033cc; TEXT-DECORATION: none
}
A:visited {
	COLOR: #0033cc; TEXT-DECORATION: none
}
A:active {
	COLOR: #ff0000; TEXT-DECORATION: none
}
A:hover {
	COLOR: #000000; TEXT-DECORATION: underline
}
.header	{
font-family: Tahoma, Verdana; font-size: 9pt; color: #FFFFFF; background-color: #00CCCC
}
.category{
font-family: Tahoma, Verdana; font-size: 9pt; color: #000000; background-color: #EFEFEF
}
</STYLE>
<title>�༭����</title>
<script language="javascript">
function submitonce(theform)
{
//if IE 4+ or NS 6+
if (document.all||document.getElementById){
//screen thru every element in the form, and hunt down "submit" and "reset"
for (i=0;i<theform.length;i++){
var tempobj=theform.elements[i]
if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
//disable em
tempobj.disabled=true
}
}
}
</script>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 onload="document.forms[0].fenlei1.disabled=true,document.forms[0].fenlei2.disabled=true,document.forms[0].teacher.disabled=true;">
<br>
<form action="editok.asp" method="post" onSubmit=submitonce(this)>
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width="480" border="1" align="center" cellpadding=1>
<tr><td align="center" class="header" colspan=2>�༭����</td></tr>
  
  <input type=hidden name="fenlei1" size=15 value="<%=fenlei1%>">
  <input type=hidden name="fenlei2" size=15 value="<%=fenlei2%>">
  <input type=hidden name="teacher" size=15 value="<%=teacher%>">
  <tr>
    <td align="center">�γ�����</td>
    <td>&nbsp;<input type=text name="course" size=15 value="<%=rs("course")%>">��Ϊ������ң������ʦע�⾡��ͳһ���ƣ�</td>
  </tr>
  <tr>
    <td align="center">���ϱ���</td>
    <td>&nbsp;<input type=text name="title" size=52 value="<%=rs("title")%>"></td>
  </tr>
  <tr>
    <td align="center">�����ϴ�</td>
    <td>
<%if session("admin") = "admin" then%>
&nbsp;����
<%else%>
&nbsp;<iframe name="up" frameborder=0 width=400 height=26 scrolling=no src=upload.asp></iframe>
<%end if%>
    </td>
  </tr>
  <tr>
    <td align="center">���ϵ�ַ</td>
    <td>&nbsp;<input type=text name="fileurl" size=50 value="<%=rs("fileurl")%>">
        <br>&nbsp;��������������������վ����������д���ӵ�ַ��
    </td>
  </tr>
  <tr>
    <td align="center">���ϴ�С</td>
    <td>&nbsp;<input type=text name="filesize" size=15 value=<%=rs("filesize")%>>K</td>
  </tr>
  <tr>
    <td align="center">��������</td>
    <td>
<%
curtypeid = rs("idoftype")
sql1 = "select * from type"
set rs1 = server.createobject("adodb.recordset")
rs1.open sql1,conn,1,1
%>
&nbsp;<select name="type">
<option value="" selected>��ѡ��</option>
<%
do while not rs1.eof
%>
 <option value="<%=rs1("typeid")%>"<%if curtypeid=rs1("typeid") then response.write " selected"%>><%=rs1("type")%></option>
<%
 rs1.movenext
loop
response.write "</select>"
rs1.close
set rs1 = nothing
%>
    </td>
  </tr>
  <tr>
    <td align="center">���ϼ��</td>
    <td>
        &nbsp;<textarea name="content" cols="40" rows="6"><%=rs("content")%></textarea>
        <input type=hidden name="id" value="<%=id%>">
    </td>
  </tr>
</table>
<center><br><input type=submit name="submit" value="�޸�">&nbsp;&nbsp;<input type=reset name="reset" value="�ָ�"></center>
</form>
</body>
</html>
<%
rs.close
set rs = nothing
conn.close
set conn = nothing
%>