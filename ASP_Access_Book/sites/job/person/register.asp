<% Response.Buffer=True %>
<!--#include file="../inc/person.inc"-->
<% uname=session("puid") 
   modify=request("modify")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from person where uname='"&uname&"'"
rs.open sql,conn,1,1 %>
<% if modify<>"ture" and rs("job") <> "" then
   response.write"<SCRIPT language=JavaScript>alert('���Ѿ���¼���˼������벻Ҫ�ظ���¼��');"
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
<title>�����˲š�&gt;��¼��ְ����</title>
<%else%>
<title>�����˲š�&gt;������ְ����</title>
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
          �� </td>
    </tr>
    <tr>
        <td width="139" valign="top" bgcolor="#53BEB0"> �� 
          <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="118" height="280">
            <tr>
              <td width="100%" height="163" background="../images/stat-bg.GIF" valign="top">
        <p align="center"><br>
        <a href="main.asp">��¼��ҳ</a><br>
        <br>
        <a href="register.asp">��¼��ְ����</a><br>
        <br>
        <a href="modify.asp">������ְ����</a><br>
        <br>
        <a href="../changepwd.asp?stype=person" target="_blank">�޸ĵ�¼����</a><br>
        <br><a href="search.asp">ȫ��ְλ�б�</a>
              </td>
            </tr>
            <tr>
              <td width="100%" height="117" background="../images/stat-bg.GIF" valign="top">
        <p align="center"><br>
        <br><a href="favorite.asp">�ҵ��ղؼ�</a><br>
        <br>
        <a href="mailbox.asp">�ҵ�����</a><br>
        <br>
        <a href="../exit.asp">�˳���¼</a>
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
    
      ��
    
        <table border="1" cellpadding="0" cellspacing="0" width="421" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
          <tr>
         <td width="417" height="18" valign="bottom" bgcolor="#C6CEDE" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
              <p align="center">&nbsp;=== ���˻������� ===</td>                            
          </tr>
          <tr>
            <td width="417" height="20" valign="bottom">
              <p align="left">&nbsp;�û���:<FONT COLOR="#FF0000"><%=uname%></FONT> &nbsp;           
            </td>
          </tr>
          <tr>
            
			<td width="417" height="20" valign="bottom" bgcolor="#F3F3F3">
              &nbsp;��ʵ����:<%if modify<>"ture" then%><input type="text" name="iname" valign="bottom" size="12" maxLength=4 value="<%=rs("iname")%>">
                <%else%>
                <%=rs("iname")%> 
                <%end if%>
                (<font color="#FF0000">*</font>����)</td>
          </tr>
          <tr>
              <td width="417" height="28"> 
                <p align="left">&nbsp;�Ա�:
                  <input type="radio"<%if modify<>"true" or rs("sex") ="��" then Response.Write "checked"%> value="��" name="sex">
                  �� 
                  <input type="radio" <%if rs("sex") ="Ů" then Response.Write "checked"%> name="sex" value="Ů">
                  Ů (<font color="#FF0000">*</font>����) </p> 
            </td>
          </tr>
          <tr>
              <td width="417" height="28" bgcolor="#F3F3F3"> 
                <p align="left">&nbsp;��������:<INPUT maxLength=10 size=10 name=bday maxLength=10 value="<%=rs("bday")%>"> 
                  (<font color="#FF0000">*</font>�������2000-1-1) </p> 
            </td>
          </tr>
          <tr>
            <td width="417" height="28">
              &nbsp;���֤����:<input type="text" name="code" size="18" maxLength=18 value="<%=rs("code")%>">
                (<font color="#FF0000">*</font>����) </td>
          </tr>
          <tr>
              <td width="417" height="28" bgcolor="#F3F3F3"> 
                <% if modify<>"ture" then %>
                <p align="left">&nbsp;����:
                  <input type="text" name="mzhu" size="12" maxLength=12 value="����">
                  (<font color="#FF0000">*</font>����) 
                  <%else%>
                <p align="left">&nbsp;����:
                  <input type="text" name="mzhu" size="12" maxLength=12 value="<%=rs("mzhu")%>">
                  (<font color="#FF0000">*</font>����) 
                  <%end if%>
                </p> 
            </td>
          </tr>
          <tr>
              <td width="417" height="28"> 
                <p align="left">&nbsp;����״��:
                  <SELECT size=1 name=marry>
                    <OPTION value=δ�� <%if rs("marry") ="δ��" then Response.Write "selected"%>>δ�� 
                    <OPTION value=�ѻ� <%if rs("marry") ="�ѻ�" then Response.Write "selected"%>>�ѻ� 
                    <OPTION value=���� <%if rs("marry") ="����" then Response.Write "selected"%>>����</OPTION>
                  </SELECT>
                  (<font color="#FF0000">*</font>����)</p>
            </td>
          </tr>
          <tr>
              <td width="417" height="28" bgcolor="#F3F3F3"> 
                <p align="left">&nbsp;�������ڵ�:
                  <SELECT size=1 name=hka>
                    <OPTION>���������б���ѡ��</OPTION>
                    <OPTION value=������ <%if rs("hka") ="������" then Response.Write "selected"%>>������ 
                    <OPTION value=����� <%if rs("hka") ="�����" then Response.Write "selected"%>>����� 
                    <OPTION value=�Ϻ��� <%if rs("hka") ="�Ϻ���" then Response.Write "selected"%>>�Ϻ��� 
                    <OPTION value=������ <%if rs("hka") ="������" then Response.Write "selected"%>>������ 
                    <OPTION value=����ʡ <%if rs("hka") ="����ʡ" then Response.Write "selected"%>>����ʡ 
                    <OPTION value=����ʡ <%if rs("hka") ="����ʡ" then Response.Write "selected"%>>����ʡ 
                    <OPTION value=����ʡ <%if rs("hka") ="����ʡ" then Response.Write "selected"%>>����ʡ 
                    <OPTION value=�㶫ʡ <%if rs("hka") ="�㶫ʡ" then Response.Write "selected"%>>�㶫ʡ 
                    <OPTION value=�Ĵ�ʡ <%if rs("hka") ="�Ĵ�ʡ" then Response.Write "selected"%>>�Ĵ�ʡ 
                    <OPTION value=����ʡ <%if rs("hka") ="����ʡ" then Response.Write "selected"%>>����ʡ 
                    <OPTION value=����ʡ <%if rs("hka") ="����ʡ" then Response.Write "selected"%>>����ʡ 
                    <OPTION value=����ʡ <%if rs("hka") ="����ʡ" then Response.Write "selected"%>>����ʡ 
                    <OPTION value=����ʡ <%if rs("hka") ="����ʡ" then Response.Write "selected"%>>����ʡ 
                    <OPTION value=����ʡ <%if rs("hka") ="����ʡ" then Response.Write "selected"%>>����ʡ 
                    <OPTION value=����ʡ <%if rs("hka") ="����ʡ" then Response.Write "selected"%>>����ʡ 
                    <OPTION value=����ʡ <%if rs("hka") ="����ʡ" then Response.Write "selected"%>>����ʡ 
                    <OPTION value=�ຣʡ <%if rs("hka") ="�ຣʡ" then Response.Write "selected"%>>�ຣʡ 
                    <OPTION value=ɽ��ʡ <%if rs("hka") ="ɽ��ʡ" then Response.Write "selected"%>>ɽ��ʡ 
                    <OPTION value=ɽ��ʡ <%if rs("hka") ="ɽ��ʡ" then Response.Write "selected"%>>ɽ��ʡ 
                    <OPTION value=����ʡ <%if rs("hka") ="����ʡ" then Response.Write "selected"%>>����ʡ 
                    <OPTION value=����ʡ <%if rs("hka") ="����ʡ" then Response.Write "selected"%>>����ʡ 
                    <OPTION value=�㽭ʡ <%if rs("hka") ="�㽭ʡ" then Response.Write "selected"%>>�㽭ʡ 
                    <OPTION value=������ʡ <%if rs("hka") ="������ʡ" then Response.Write "selected"%>>������ʡ 
                    
                  </SELECT>
                  (<font color="#FF0000">*</font>����)</p>
            </td>
          </tr>
          <tr>
              <td width="417" height="28"> 
                <p align="left"> &nbsp;����ߵĽ����̶�:
                  <SELECT size=1 
                  name=edu>
                    <OPTION>���������б���ѡ��</OPTION>
                    <OPTION value=���� <%if rs("edu") ="����" then Response.Write "selected"%>>���� 
                    <OPTION value=���� <%if rs("edu") ="����" then Response.Write "selected"%>>���� 
                    <OPTION value=�м� <%if rs("edu") ="�м�" then Response.Write "selected"%>>�м� 
                    <OPTION value=��ר <%if rs("edu") ="��ר" then Response.Write "selected"%>>��ר 
                    <OPTION value=��ר <%if rs("edu") ="��ר" then Response.Write "selected"%>>��ר 
                    <OPTION value=���� <%if rs("edu") ="����" then Response.Write "selected"%>>���� 
                    <OPTION value=˶ʿ <%if rs("edu") ="˶ʿ" then Response.Write "selected"%>>˶ʿ 
                    <OPTION value=��ʿ <%if rs("edu") ="��ʿ" then Response.Write "selected"%>>��ʿ</OPTION>
                  </SELECT>
                  (<font color="#FF0000">*</font>����)</p> 
            </td>
          </tr>
          <tr>
            <td width="417" height="28" bgcolor="#F3F3F3">
              <p align="left"> 
              &nbsp;ר ҵ:<INPUT maxLength=60 size=30                            
                  name=zye maxLength=30 value="<%=rs("zye")%>">(<font color="#FF0000">*</font>����)</p> 
            </td>
          </tr>
          <tr>
            <td width="417" height="28">
              <p align="left"> 
              &nbsp;��ҵԺУ:<INPUT maxLength=60 size=40 
                  name=school maxLength=40 value="<%=rs("school")%>">(<font color="#FF0000">*</font>����)</p> 
            </td>
          </tr>
          <tr>
              <td width="417" height="28" bgcolor="#F3F3F3"> 
                <p align="left"> &nbsp;������ò:
                  <SELECT size=1 name=zzmm>
                    <OPTION>���������б���ѡ��</OPTION>
                    <OPTION value=��Ա <%if rs("zzmm") ="��Ա" then Response.Write "selected"%>>��Ա 
                    <OPTION value=��Ա <%if rs("zzmm") ="��Ա" then Response.Write "selected"%>>��Ա 
                    <OPTION value=Ⱥ�� <%if rs("zzmm") ="Ⱥ��" then Response.Write "selected"%>>Ⱥ�� 
                    <OPTION value=�������� <%if rs("zzmm") ="��������" then Response.Write "selected"%>>�������� 
                    <OPTION value=���� <%if rs("zzmm") ="����" then Response.Write "selected"%>>����</OPTION>
                  </SELECT>
                  (<font color="#FF0000">*</font>����)</p> 
            </td>
          </tr>
          <tr>
              <td width="417" height="28"> 
                <p align="left"> &nbsp;����ְ��:
                  <SELECT size=1 name=zchen>
                    <OPTION>���������б���ѡ��</OPTION>
                    <OPTION value=�߼� <%if rs("zchen") ="�߼�" then Response.Write "selected"%>>�߼� 
                    <OPTION value=�м� <%if rs("zchen") ="�м�" then Response.Write "selected"%>>�м� 
                    <OPTION value=���� <%if rs("zchen") ="����" then Response.Write "selected"%>>���� 
                    <OPTION value=���� <%if rs("zchen") ="����" then Response.Write "selected"%>>���� 
                    <OPTION value=���� <%if rs("zchen") ="����" then Response.Write "selected"%>>����</OPTION>
                  </SELECT>
                  (<font color="#FF0000">*</font>����)</p> 
            </td>
          </tr>
          <tr>
            <td valign="top" width="417">
              
              <p align="center"> 
              <br>
              <% if modify<>"ture" then %>
			  <input type="button" value="��һ��" onClick="check()"></p>  
              <%else%>
              <input type="button" value="�� ��" onClick="checkmod()">  
             <br>
             <%end if%>      
            <br>
            </td>
          </tr>
        </table>
      </td>
  <center>
      <td width="1" height="243" valign="top" bgcolor="#00006A"></td>
      <td width="138" height="243" valign="top" bgcolor="#F3F3F3">��</td>
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