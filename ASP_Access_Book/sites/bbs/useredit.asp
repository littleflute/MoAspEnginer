<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<!--#include file="check.asp"-->
<%
Dim rs__MMColParam
If (Request.QueryString("id") <> "") Then 
  rs__MMColParam = Request.QueryString("id")
  '�ж��Ƿ�����޸�Ȩ��
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
<title>�û��޸�����</title>
<!-- #include file="an.htm" -->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid; background-color: #e3f1d1;">
  <tr> 
    <td width="468" height="24"><a href="index.asp">��̳��̳</a>���޸�����</td>
    <td width="290">|<a href="uerlist.asp">�û��б�</a> | <a href="adduer.asp">�û�ע��</a>| 
      <a href="find.asp">���һ�Ա</a>| <a href="uerlist.asp?id=2">��������</a>|</td>
  </tr>
</table>
  <table width="760" border="1" align="center" cellpadding="0" cellspacing="1" bordercolor="#CCCCCC">
<form action="usereditsave.asp" method="post" name="form" id="form">
    <tr > 
    <td height="21" colspan="2" background="pic/h3.gif"><div align="center"><strong><font color="#FFFFFF">�޸�����</font></strong></div></td>
  </tr>
    <tr> 
      <td colspan="2" bgcolor="#FFFFFF" class="p9"><%=(rs.Fields.Item("id").Value)%>-�޸�����</td>
    </tr>
    <tr> 
      <td colspan="2"> <div align="right"></div>
        <input name="id" type="hidden" id="id" value="<%=(rs.Fields.Item("id").Value)%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <p align="right">*�� �룺</td>
      <td width="432"> <input name="psw" type="password" id="psw" value="<%=(rs.Fields.Item("psw").Value)%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <div align="right">��������һ�����룺</div></td>
      <td width="432"> <input name="pswc" type="password" id="psw0" value="<%=(rs.Fields.Item("psw").Value)%>" size="20"></td>
    </tr>
    <tr> 
      <td height="18" bgcolor="#e3f1d1" width="261"> <div align="right">*�� ��</div></td>
      <td width="432"><input <%If (CStr((rs.Fields.Item("sex").Value)) = CStr("boy")) Then Response.Write("CHECKED") : Response.Write("")%> name="sex" type="radio" value="boy" checked>
        ˧�� &nbsp; <input <%If (CStr((rs.Fields.Item("sex").Value)) = CStr("girl")) Then Response.Write("CHECKED") : Response.Write("")%> type="radio" name="sex" value="girl">
        ��Ů</td>
    </tr>
    <tr> 
      <td height="30" bgcolor="#e3f1d1"> <div align="right">���Ի�ͷ��</div></td>
      <td><select class="t2" onChange="document.images['idface'].src=options[selectedIndex].value;" size="1" name="face">
          <option value="images/01.gif">�û�ͷ��-01 </option>
          <option value="images/02.gif">�û�ͷ��-02 </option>
          <option value="images/03.gif">�û�ͷ��-03 </option>
          <option value="images/04.gif">�û�ͷ��-04 </option>
          <option value="images/05.gif">�û�ͷ��-05 </option>
          <option value="images/06.gif">�û�ͷ��-06 </option>
          <option value="images/07.gif">�û�ͷ��-07 </option>
          <option value="images/08.gif">�û�ͷ��-08 </option>
          <option value="images/09.gif">�û�ͷ��-09 </option>
          <option value="images/10.gif">�û�ͷ��-10 </option>
          <option value="images/11.gif">�û�ͷ��-11 </option>
          <option value="images/12.gif">�û�ͷ��-12 </option>
          <option value="images/13.gif">�û�ͷ��-13 </option>
          <option value="images/14.gif">�û�ͷ��-14 </option>
          <option value="images/15.gif">�û�ͷ��-15 </option>
          <option value="images/16.gif">�û�ͷ��-16 </option>
          <option value="images/17.gif">�û�ͷ��-17 </option>
          <option value="images/18.gif">�û�ͷ��-18 </option>
          <option value="images/19.gif">�û�ͷ��-19 </option>
          <option value="images/20.gif">�û�ͷ��-20</option>
        </select> <img src="<%=rs("face")%>" alt="�����������" align="absmiddle" class="t2" id="idface"><font class=j1>[<a 
      href="tx.htm" target=_blank>����ͷ��</a>]</font></td>
    </tr>
    <tr> 
      <td height="19" bgcolor="#e3f1d1"> <div align="right">ͷ�Σ�</div></td>
      <td><input name="head" type="text" id="head" value="<%=(rs.Fields.Item("head").Value)%>"></td>
    </tr>
    <tr> 
      <td height="19" bgcolor="#e3f1d1" width="261"> <div align="right">Q Q��</div></td>
      <td width="432"><input name="qq" type="text" id="qq" value="<%=(rs.Fields.Item("qq").Value)%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <div align="right">*E-mail��</div></td>
      <td width="432"><input name="email" type="text" id="email" value="<%=(rs.Fields.Item("email").Value)%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <p align="right">������ҳ��</td>
      <td width="432"><input name="homepage" type="text" id="homepage" value="<%=(rs.Fields.Item("homepage").Value)%>"> 
      </td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <div align="right">Ժϵ&amp;�꼶��</div></td>
      <td width="432"> <input name="college" type="text" id="college" value="<%=(rs.Fields.Item("college").Value)%>"> 
      </td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <div align="right">*�Ƿ�Э��Ա��</div></td>
      <td width="432"> <input <%If (CStr((rs.Fields.Item("huiyuan").Value)) = CStr("yes")) Then Response.Write("CHECKED") : Response.Write("")%> name="huiyuan" type="radio" value="yes" checked>
        �� &nbsp; <input <%If (CStr((rs.Fields.Item("huiyuan").Value)) = CStr("no")) Then Response.Write("CHECKED") : Response.Write("")%> type="radio" name="huiyuan" value="no">
        ��</td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1">&nbsp;</td>
      <td bgcolor="#e3f1d1">&nbsp;</td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <div align="right">*������</div></td>
      <td width="432"><input name="name" type="text" id="name" value="<%=(rs.Fields.Item("name").Value)%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <div align="right">���ţ�</div></td>
      <td width="432"> <select name="part" id="part">
          <option value="�ǻ�Ա" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("�ǻ�Ա" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>�ǻ�Ա</option>
          <option value="�᳤" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("�᳤" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>�᳤</option>
          <option value="���᳤" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("���᳤" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>���᳤</option>
          <option value="���鴦" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("���鴦" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>���鴦</option>
          <option value="��Ŀ��" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("��Ŀ��" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>��Ŀ��</option>
          <option value="���ز�" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("���ز�" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>���ز�</option>
          <option value="������" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("������" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>������</option>
          <option value="OEC���" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("OEC���" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>OEC���</option>
          <option value="��Ա���ֲ�" <%If (Not isNull((rs.Fields.Item("part").Value))) Then If ("��Ա���ֲ�" = CStr((rs.Fields.Item("part").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>��Ա���ֲ�</option>
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
      <td bgcolor="#e3f1d1" width="261"> <div align="right">ְ��</div></td>
      <td width="432"> <select name="class" id="class">
          <option value="�ǻ�Ա" <%If (Not isNull((rs.Fields.Item("class").Value))) Then If ("�ǻ�Ա" = CStr((rs.Fields.Item("class").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>�ǻ�Ա</option>
          <option value="����" <%If (Not isNull((rs.Fields.Item("class").Value))) Then If ("����" = CStr((rs.Fields.Item("class").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>����</option>
          <option value="������" <%If (Not isNull((rs.Fields.Item("class").Value))) Then If ("������" = CStr((rs.Fields.Item("class").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>������</option>
          <option value="��Ա" <%If (Not isNull((rs.Fields.Item("class").Value))) Then If ("��Ա" = CStr((rs.Fields.Item("class").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>��Ա</option>
          <option value="�᳤" <%If (Not isNull((rs.Fields.Item("class").Value))) Then If ("�᳤" = CStr((rs.Fields.Item("class").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>�᳤</option>
          <option value="���᳤" <%If (Not isNull((rs.Fields.Item("class").Value))) Then If ("���᳤" = CStr((rs.Fields.Item("class").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>���᳤</option>
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
      <td bgcolor="#e3f1d1" width="261"> <div align="right">�� ����</div></td>
      <td width="432"> <input name="tel" type="text" id="tel" value="<%=(rs.Fields.Item("tel").Value)%>"></td>
    </tr>
    <tr> 
      <td bgcolor="#e3f1d1" width="261"> <div align="right">סַ��</div></td>
      <td width="432"> <input name="addr" type="text" id="addr" value="<%=(rs.Fields.Item("addr").Value)%>"></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center"><br>
          <input type="submit" name="Submit" value="ȷ �� �� ��">
        </div></td>
    </tr>
</form>  </table>
<%
rs.Close()
Set rs = Nothing
%>
<!--#include file="foot.asp"-->