<% Response.Buffer=True %>
<!--#include file="../inc/company.inc"-->
<!--#include file="../inc/html.inc"-->
<% uname=session("cuid")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from company where uname='"&uname&"'"
rs.open sql,conn,1,1 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/register.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�����˲š�&gt;��¼��˾����</title>
</head>
<SCRIPT language=JavaScript src="../inc/validate.js"></SCRIPT>
<SCRIPT language=JavaScript src="../inc/companyreg.js"></SCRIPT>
<FORM name=register action=register.asp method=post>
<body topmargin="0" leftmargin="0">
<!--#include file="../inc/top2.htm"--> 
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="61">
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
		<a href="register.asp">��¼/���¹�˾����</a><br>
        <br>
		<a href="publish.asp">����/������Ƹ��Ϣ</a><br>
        <br>
        <a href="../changepwd.asp?stype=company" target="_blank">�޸ĵ�¼����</a><br>
        <br>
        <a href="search.asp">ȫ���˲��б�</a>
              </td>
            </tr>
            <tr>
              <td width="100%" height="117" background="../images/stat-bg.GIF" valign="top">
        <p align="center"><br>
        <br>
        <a href="favorite.asp">�ҵ��ղؼ�<br>
        <br>
        <a href="mailbox.asp">�ҵ�����</a><br>
        <br>
        <a href="../exit.asp">�˳���¼</a>
              </td>
            </tr>
          </table>
        </div>
        <p align="center"> </td>
      <td width="31" height="325" valign="top"><img border="0" src="../images/selfk.GIF"></td>
  </center>
      <td width="481" height="325">
        ��
        <div align="left">
        <table border="1" cellpadding="0" cellspacing="0" width="94%" height="263" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
          <tr>
            <td width="100%" height="18" valign="bottom" bgcolor="#C6CEDE" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
              <p align="center">=== ��˾��ϸ���� ===</td>                           
          </tr>
          <tr>
            <td width="100%" height="20" valign="bottom">
              &nbsp;�û���:<FONT COLOR="#FF0000"><%=uname%></FONT> &nbsp;</td>                           
          </tr>
          <tr>
            
			<td width="100%" height="22" valign="bottom" bgcolor="#F3F3F3">
            &nbsp;��˾����:<%if rs("cname")<>"" then%>
			<%=rs("cname")%><%else%>    
			<input type="text" name="cname" size="40" maxLength=40 value="<%=rs("cname")%>">    
			<%end if%>    
			</td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom">
              <p align="left">&nbsp;������ҵ:<SELECT size=1 name=trade> 
              <OPTION>���������б���ѡ��</OPTION> <OPTION 
              value=�����ҵ <%if rs("trade") ="�����ҵ" then Response.Write "selected"%>>�����ҵ</OPTION> <OPTION value=���ӵ�����ͨѶ�豸 <%if rs("trade") ="���ӵ�����ͨѶ�豸" then Response.Write "selected"%>>���ӵ�����ͨѶ�豸</OPTION> 
              <OPTION value=������ơ��Ƽ�����  <%if rs("trade") ="������ơ��Ƽ����� " then Response.Write "selected"%>>������ơ��Ƽ�����</OPTION> <OPTION 
              value=�����豸������ <%if rs("trade") ="�����豸������ " then Response.Write "selected"%>>�����豸������</OPTION> <OPTION value=�����Ǳ��������� <%if rs("trade") ="�����Ǳ���������" then Response.Write "selected"%>>�����Ǳ���������</OPTION> 
              <OPTION value=��е���켰�豸 <%if rs("trade") ="��е���켰�豸" then Response.Write "selected"%>>��е���켰�豸</OPTION> <OPTION 
              value=���ֳ������������� <%if rs("trade") ="���ֳ�������������" then Response.Write "selected"%>>���ֳ�������������</OPTION> <OPTION value=���������� <%if rs("trade") ="����������" then Response.Write "selected"%>>����������</OPTION> 
              <OPTION value=��ѧ������������ҩ <%if rs("trade") ="��ѧ������������ҩ" then Response.Write "selected"%>>��ѧ������������ҩ</OPTION> <OPTION 
              value=�������������Ʒ <%if rs("trade") ="�������������Ʒ" then Response.Write "selected"%>>�������������Ʒ</OPTION> <OPTION value=��֯����װ <%if rs("trade") ="��֯����װ" then Response.Write "selected"%>>��֯����װ</OPTION> 
              <OPTION value=ũ�������� <%if rs("trade") ="ũ��������" then Response.Write "selected"%>>ũ��������</OPTION> <OPTION value=�Ṥ <%if rs("trade") ="�Ṥ" then Response.Write "selected"%>>�Ṥ</OPTION> 
              <OPTION value=���ز���������װ�� <%if rs("trade") ="���ز���������װ��" then Response.Write "selected"%>>���ز���������װ��</OPTION> <OPTION 
              value=��ֽ��ӡˢ����װ <%if rs("trade") ="��ֽ��ӡˢ����װ" then Response.Write "selected"%>>��ֽ��ӡˢ����װ</OPTION> <OPTION value=ľ�ġ��Ҿ� <%if rs("trade") ="ľ�ġ��Ҿ�" then Response.Write "selected"%>>ľ�ġ��Ҿ�</OPTION> 
              <OPTION value=��桢�߻������ <%if rs("trade") ="��桢�߻������" then Response.Write "selected"%>>��桢�߻������</OPTION> <OPTION 
              value=���ų��桢�㲥���� <%if rs("trade") ="���ų��桢�㲥����" then Response.Write "selected"%>>���ų��桢�㲥����</OPTION> <OPTION 
              value=��Ϣ��ѯ���˲Ž��� <%if rs("trade") ="��Ϣ��ѯ���˲Ž���" then Response.Write "selected"%>>��Ϣ��ѯ���˲Ž���</OPTION> <OPTION 
              value=���ڡ����� <%if rs("trade") ="���ڡ�����" then Response.Write "selected"%>>���ڡ�����</OPTION> <OPTION value=��ͨ���� <%if rs("trade") ="��ͨ����" then Response.Write "selected"%>>��ͨ����</OPTION> <OPTION 
              value=���������� <%if rs("trade") ="����������" then Response.Write "selected"%>>����������</OPTION> <OPTION value=�ۺ��Թ�����ҵ <%if rs("trade") ="�ۺ��Թ�����ҵ" then Response.Write "selected"%>>�ۺ��Թ�����ҵ</OPTION> 
              <OPTION value=������������ҵ <%if rs("trade") ="������������ҵ" then Response.Write "selected"%>>������������ҵ</OPTION> <OPTION 
              value=������ҵ <%if rs("trade") ="������ҵ" then Response.Write "selected"%>>������ҵ</OPTION> <OPTION value=ҽ�ơ�������ҵ <%if rs("trade") ="ҽ�ơ�������ҵ" then Response.Write "selected"%>>ҽ�ơ�������ҵ</OPTION> 
              <OPTION value=�Ļ������� <%if rs("trade") ="�Ļ�������" then Response.Write "selected"%>>�Ļ�������</OPTION> <OPTION value=�����������칫��Ʒ <%if rs("trade") ="�����������칫��Ʒ" then Response.Write "selected"%>>�����������칫��Ʒ</OPTION> <OPTION value=����������� <%if rs("trade") ="�����������" then Response.Write "selected"%>>�����������</OPTION> 
              <OPTION value=��ҵ���� <%if rs("trade") ="��ҵ����" then Response.Write "selected"%>>��ҵ����</OPTION> <OPTION value=���ʹ��� <%if rs("trade") ="���ʹ���" then Response.Write "selected"%>>���ʹ���</OPTION> 
              <OPTION value=��ʳ������ҵ <%if rs("trade") ="��ʳ������ҵ" then Response.Write "selected"%>>��ʳ������ҵ</OPTION> <OPTION 
              value=���͡���ʳ <%if rs("trade") ="���͡���ʳ" then Response.Write "selected"%>>���͡���ʳ</OPTION></SELECT>
              </p>
            </td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom" bgcolor="#F3F3F3">
              &nbsp;��ҵ����:<SELECT size=1 name=cxz> <OPTION 
              value="" selected>���������б���ѡ��</OPTION> <OPTION 
              value=�������� <%if rs("cxz") ="��������" then Response.Write "selected"%>>��������</OPTION> <OPTION value=������� <%if rs("cxz") ="�������" then Response.Write "selected"%>>�������</OPTION> <OPTION 
              value=��ҵ��λ <%if rs("cxz") ="��ҵ��λ" then Response.Write "selected"%>>��ҵ��λ</OPTION> <OPTION value=������ҵ <%if rs("cxz") ="������ҵ" then Response.Write "selected"%>>������ҵ</OPTION> <OPTION 
              value=������ҵ <%if rs("cxz") ="������ҵ" then Response.Write "selected"%>>������ҵ</OPTION> <OPTION value=������ҵ <%if rs("cxz") ="������ҵ" then Response.Write "selected"%>>������ҵ</OPTION> <OPTION value=������ҵ <%if rs("cxz") ="������ҵ" then Response.Write "selected"%>>������ҵ</OPTION> <OPTION 
              value=˽Ӫ��ҵ <%if rs("cxz") ="˽Ӫ��ҵ" then Response.Write "selected"%>>˽Ӫ��ҵ</OPTION> </SELECT>
            </td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom">
              &nbsp;��������:<B><SELECT size=1 name=area> 
              <OPTION>���������б���ѡ��</OPTION>
			  <OPTION value=������ <%if rs("area") ="������" then Response.Write "selected"%>>������
              <OPTION value=����� <%if rs("area") ="�����" then Response.Write "selected"%>>�����
              <OPTION value=�Ϻ��� <%if rs("area") ="�Ϻ���" then Response.Write "selected"%>>�Ϻ���
			  <OPTION value=������ <%if rs("area") ="������" then Response.Write "selected"%>>������
              <OPTION value=����ʡ <%if rs("area") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=����ʡ <%if rs("area") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=����ʡ <%if rs("area") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=�㶫ʡ <%if rs("area") ="�㶫ʡ" then Response.Write "selected"%>>�㶫ʡ
			  <OPTION value=�Ĵ�ʡ <%if rs("area") ="�Ĵ�ʡ" then Response.Write "selected"%>>�Ĵ�ʡ
              <OPTION value=����ʡ <%if rs("area") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=����ʡ <%if rs("area") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=����ʡ <%if rs("area") ="����ʡ" then Response.Write "selected"%>>����ʡ
			  <OPTION value=����ʡ <%if rs("area") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=����ʡ <%if rs("area") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=����ʡ <%if rs("area") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=����ʡ <%if rs("area") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=�ຣʡ <%if rs("area") ="�ຣʡ" then Response.Write "selected"%>>�ຣʡ
			  <OPTION value=ɽ��ʡ <%if rs("area") ="ɽ��ʡ" then Response.Write "selected"%>>ɽ��ʡ
              <OPTION value=ɽ��ʡ <%if rs("area") ="ɽ��ʡ" then Response.Write "selected"%>>ɽ��ʡ
			  <OPTION value=����ʡ <%if rs("area") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=����ʡ <%if rs("area") ="����ʡ" then Response.Write "selected"%>>����ʡ
			  <OPTION value=�㽭ʡ <%if rs("area") ="�㽭ʡ" then Response.Write "selected"%>>�㽭ʡ
			  <OPTION value=������ʡ <%if rs("area") ="������ʡ" then Response.Write "selected"%>>������ʡ

              </SELECT></B>
            </td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom" bgcolor="#F3F3F3">
              &nbsp;��������:<input type="text" name="fdate" size="10" maxLength=10 value="<%=rs("fdate")%>">
            </td>                 
          </tr>
          <tr>
            <% kfund=rs("fund")
			  if kfund="δ֪" then kfund="" end if %>
            <td width="100%" height="28" valign="bottom">
              &nbsp;ע���ʽ�;<input type="text" name="fund" size="6" maxLength=6 value="<%=kfund%>">�������</td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom" bgcolor="#F3F3F3">
              &nbsp;ͨѶ��ַ:<input type="text" name="address" size="30" maxLength=40 value="<%=rs("address")%>">
            </td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom">
              &nbsp;��������;<input type="text" name="zip" size="6" size="6" maxLength=6 value="<%=rs("zip")%>">
            </td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom" bgcolor="#F3F3F3">
              <p align="left">&nbsp;��ϵ��:&nbsp; <input type="text" name="pname" size="20" maxLength=20 value="<%=rs("pname")%>">    
              </p></td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom">
              &nbsp;��ϵ�绰:<input type="text" name="phone" size="20" maxLength=40 value="<%=rs("phone")%>">
            </td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom" bgcolor="#F3F3F3">
              <p align="left">&nbsp;�������:<input type="text" name="fax" size="20" maxLength=40 value="<%=rs("fax")%>">
              </p></td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom">
              &nbsp;��������:<input type="text" name="email" size="20" maxLength=40 value="<%=rs("email")%>">
            </td>                 
          </tr>
          <tr>
          <td width="100%" height="28" valign="bottom" bgcolor="#F3F3F3">
           &nbsp;��˾��վ:<input type="text" name="http" size="20" maxLength=40 
		   <% if rs("http") <>"" then%> value="<%=rs("http")%>" <%else%> value="http://"<%end if%>>
            </td>                 
          </tr>
          <tr>
            <td width="100%" valign="top">
              <p align="center"><br>
              ��˾���:<br>
              <br>
              <%if rs("cname")<>"" then 
			  kjianj=replace(rs("jianj"),"<br>",chr(13))
              kjianj=replace(kjianj,"&nbsp;"," ")
			  else
			  kjianj="" %>
			  <%end if%>      
   <textarea rows="9" name="jianj" cols="62" style="background-color: #F3F3F3; color: #000060"><%=kjianj%></textarea>    
              </p>
              <p align="center"><input type="button" value="ȷ ��" <%if rs("cname")<>"" then%>onClick="checkmod()"<%else%>onClick="check()"<%end if%>>
              <br>
              <br>
              </p></td>
          </tr>
        </table>
        </div>
      </td>
  <center>
      <td width="1" height="325" valign="top" bgcolor="#00006A"></td>
      <td width="122" height="325" valign="top" bgcolor="#F3F3F3">��</td>
    </tr>
    <tr>
      <td width="778" height="1" valign="top" colspan="5" bgcolor="#53BEB0"> 
        <p align="center"> </td>
    </tr>
    <tr>
      <td width="778" height="3" valign="top" colspan="5">
      <p align="center">
      </td>
    </tr>
    <tr>
      <td width="778" height="7" valign="top" colspan="5">
      <p align="center"><script language="javascript" src="../inc/copyright.js"></script>
      </td>
    </tr>
    <tr>
      <td width="778" height="3" valign="top" colspan="5">
      </td>
    </tr>
  </table>
  </center>
</div>
</form>
</body>

</html>

<% pname=request("pname") 
if pname="" then Response.End
trade=request("trade")
cxz=request("cxz")
area=request("area")
fdate=request("fdate")
fund=request("fund")
cname=request("cname")
address=request("address")
zip=request("zip")
phone=request("phone")
fax=request("fax")
email=request("email")
http=request("http")
jianj=htmlencode2(request("jianj"))
if fund="" then fund="δ֪" end if
if cname="" then cname=rs("cname") end if
if fax="" then fax="δ֪" end if
if http="" then http="http://" end if
rs.close
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from company where uname='"&uname&"'"
rs.open sql,conn,3,3
rs("trade")=trade
rs("cxz")=cxz
rs("area")=area
rs("fdate")=fdate
rs("cname")=cname
rs("fund")=fund
rs("pname")=pname
rs("address")=address
rs("zip")=zip
rs("phone")=phone
rs("fax")=fax
rs("email")=email
rs("http")=http
rs("jianj")=jianj
rs.update
rs.close
Response.Redirect "main.asp" %>