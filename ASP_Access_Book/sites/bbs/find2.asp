<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<!--#include file="conn.asp" -->

<%
Dim rs__MMColParam
dim sel
sel=Request.Form("select")
if Request.Form("name") = "" then
errmsg="����дҪ���ҵ���Ϣ"
end if

if errmsg<>"" then
    response.write("<script>alert('" & errmsg & "');history.go(-1)</script>")
    response.end
end if


'rs__MMColParam = "1"
If (Request.Form("name") <> "") Then 
  rs__MMColParam = Request.Form("name")
End If
%>
<%
Dim rs
Dim rs_numRows

Set rs = Server.CreateObject("ADODB.Recordset")
if sel="1" then
rs.Source ="SELECT * FROM uers WHERE name LIKE '%" + Replace(rs__MMColParam, "'", "''") + "%'"
else
rs.Source ="SELECT * FROM uers WHERE id LIKE '%" + Replace(rs__MMColParam, "'", "''") + "%'"
end if
rs.Open rs.Source,conn,1,1
rs_numRows = 0
%>
<html>
<head>
<title>Untitled Document</title>
<!-- #include file="an.htm" -->

<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid; background-color: #e3f1d1;">
  <tr> 
    <td width="475" height="24"><a href="index.asp">��̳��̳</a>��<a href="find.asp">���һ�Ա</a></td>
    <td width="283">|<a href="uerlist.asp">�û��б�</a> | <a href="adduer.asp">�û�ע��</a>| 
      <a href="find.asp">���һ�Ա</a>| <a href="uerlist.asp?id=2">��������</a>|</td>
  </tr>
  <tr align="center"> 
    <td height="21" colspan="9" background="pic/h3.gif" class="p9"><strong><font color="#FFFFFF">���һ�Ա</font></strong></td>
  </tr>
</table>
 <%
    If not rs.eof   Then
	do while not rs.eof  
  %>
<table width="760" border="1" align="center" cellpadding="0" cellspacing="1" bordercolor="#CCCCCC">
        <tr> 
          <td colspan="2" background="pic/bg_re.gif" class="pbutton"><%=(rs.Fields.Item("id").Value)%> 
            ����ϸ�������£�</td>
        </tr>
        <tr class="p9"> 
          <td width="31%" bgcolor="#e3f1d1"> <div align="right">�û�����</div></td>
          <td width="69%"><%=(rs.Fields.Item("id").Value)%></td>
        </tr>
        <tr class="p9"> 
          <td height="19" bgcolor="#e3f1d1"><div align="right">���Ի�ͷ��</div></td>
          <td><img src="<%=(rs.Fields.Item("face").Value)%>"></td>
        </tr>
        <tr class="p9"> 
          <td height="19" bgcolor="#e3f1d1"><div align="right">ͷ�Σ�</div></td>
          <td><%=(rs.Fields.Item("head").Value)%></td>
        </tr>
        <tr class="p9"> 
          <td height="19" bgcolor="#e3f1d1"><div align="right">�ȼ���</div></td>
          <td><%=(rs.Fields.Item("dengji").Value)%></td>
        </tr>
        <tr class="p9"> 
          <td height="19" bgcolor="#e3f1d1"><div align="right">����������</div></td>
          <td><%=(rs.Fields.Item("totle").Value)%></td>
        </tr>
        <tr class="p9"> 
          <td height="19" bgcolor="#e3f1d1"><div align="right">����ʱ�䣺</div></td>
          <td><%=(rs.Fields.Item("time").Value)%></td>
        </tr>
        <tr class="p9"> 
          <td height="19" bgcolor="#e3f1d1"> <div align="right">Q Q��</div></td>
          <td><%=(rs.Fields.Item("qq").Value)%></td>
        </tr>
        <tr class="p9"> 
          <td bgcolor="#e3f1d1"> <div align="right">E-mail��</div></td>
          <td><%=(rs.Fields.Item("email").Value)%></td>
        </tr>
        <tr class="p9"> 
          <td bgcolor="#e3f1d1"> <div align="right">������ҳ��</div></td>
          <td><a href="<%=(rs.Fields.Item("homepage").Value)%>"><%=(rs.Fields.Item("homepage").Value)%></a> 
          </td>
        </tr>
        <tr class="p9"> 
          <td bgcolor="#e3f1d1"> <div align="right">�Ƿ�Э��Ա��</div></td>
          <td> 
            <% if rs.Fields.Item("huiyuan").Value="yes"  then%>
            �� 
            <% Else %>
            �� 
            <% End If %>
          </td>
        </tr>
        <tr class="p9"> 
          <td bgcolor="#e3f1d1">&nbsp;</td>
          <td bgcolor="#e3f1d1">&nbsp;</td>
        </tr>
        <tr class="p9"> 
          <td bgcolor="#e3f1d1"> <div align="right">������</div></td>
          <td><%=(rs.Fields.Item("name").Value)%></td>
        </tr>
        <tr class="p9"> 
          <td bgcolor="#e3f1d1"> <div align="right">���ţ�</div></td>
          <td><%=(rs.Fields.Item("part").Value)%> </td>
        </tr>
        <tr class="p9"> 
          <td bgcolor="#e3f1d1"> <div align="right">ְ��</div></td>
          <td><%=(rs.Fields.Item("class").Value)%> </td>
        </tr>
        <tr class="p9"> 
          <td bgcolor="#e3f1d1"> <div align="right">Ժϵ&amp;�꼶��</div></td>
          <td><%=(rs.Fields.Item("college").Value)%> </td>
        </tr>
        <tr class="p9"> 
          <td bgcolor="#e3f1d1"> <div align="right">�� ����</div></td>
          <td><%=(rs.Fields.Item("tel").Value)%> </td>
        </tr>
        <tr class="p9"> 
          <td bgcolor="#e3f1d1"> <div align="right">סַ��</div></td>
          <td><%=(rs.Fields.Item("addr").Value)%> </td>
        </tr>
        <tr align="center" class="p9"> 
          <td colspan="2" bgcolor="#e3f1d1">�� <a href="javascript:history.back(1)">�� 
            ��</a> ��</td>
        </tr>
      </table>
<%
   rs.movenext
   loop  
   Else
%>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="19">&nbsp;</td>
  </tr>
  <tr> 
    <td height="22" align="center"><span class="p9">û����Ҫ�ҵļ�¼������㲻��ȷ������������Ը��ý��ٵ���</span></td>
  </tr>
  <tr>
    <td height="17" align="center">&nbsp;</td>
  </tr>
  <tr> 
    <td height="17" align="center"><p> <a href="find.asp" class="p9">����������¼���</a> 
      </p></td>
  </tr>
</table>
<% End If %>
<%
rs.Close()
Set rs = Nothing
%>
