<!--#include file="conn.asp"-->
<%
'�Ѷϴ��û��Ƿ��Ѿ����
if session("user_id")="" then
     response.redirect "notreg.htm"
	 response.end
end if

Set rs_lar = Server.CreateObject("ADODB.Recordset")
sql="select * from larchives where user_id=" & session("user_id")
rs_lar.open sql,conn,3,2
if rs_lar.eof and rs_lar.bof then
   response.redirect "notregist.htm"
   response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Namo WebEditor v4.0(Trial)">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>��������</title>
<style>
<!--
-->
</style>
<link rel="stylesheet" href="ysb.css">
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" leftmargin="5" topmargin="0" >
<table width="772" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="images/top.gif" width="772" height="112" usemap="#Map" border="0"></td>
  </tr>
</table>
<map name="FPMap1"> <area shape="rect" coords="381, 4, 446, 18" href="sendphoto.asp"> 
  <area shape="rect" coords="283, 4, 357, 18" href="register.asp"
 target="_blank"> <area shape="rect" coords="190, 4, 264, 18" href="list.asp"> 
  <area shape="rect" coords="8, 2, 71, 16" href="your.asp"> <area shape="rect" coords="102, 4, 168, 18" href="no.asp"> 
</map> 
<table border="0" width="776" cellspacing="0" cellpadding="0">
  <tr>
    <td width="3%">&nbsp; </td>
    <td width="71%">&nbsp; </td>
    <td width="1%">&nbsp; </td>
    <td width="25%"></td>
  </tr>
  <tr>
    <td width="3%" valign="top">
    </td>
    <td width="71%" valign="top"> 
      <form action="eadd.asp" method="POST">
<div align="left">
<table border="0" width="526" cellspacing="0" cellpadding="0">
    <td width="600">
      <table border="0" width="552" cellspacing="0" height="204" cellpadding="0">
                  <tr bgcolor="#067B0F"> 
                    <td height="11" colspan="3"><font color="#FFFFFF"><b>��������д���ĸ�������</b></font></td>
        </tr>
        <tr>
                    <td height="4" colspan="3"><font color="#FFFFFF">&nbsp;</font></td>
        </tr>
        <tr>
                    <td height="7" style="border-bottom-style: solid" bordercolor="#000000" colspan="3"><b><font color="#FF9900">���Ļ�������</font></b></td>
        </tr>
        <tr>
          <td width="119" height="18"><font color="#008000"><b>*��ʵ����:</b></font></td>
                    <td width="260" height="18"> 
                      <input type="text" name="name" value="<%=rs_lar("name")%>" size="10" style="border-style: solid; border-width: 1"></td>
                    <td width="173" height="18"><i>2-10�ַ�</i></td>
        </tr>
        <tr>
          <td width="119" height="18"><font color="#008000"><b>*�����Ա�:</b></font></td>
                    <td width="260" height="18"> 
                      <INPUT <%if rs_lar("sex")="��" then%>CHECKED<%end if%> name=sex type=radio
            value=��> <SPAN class=tt4>��&nbsp; <INPUT <%if rs_lar("sex")="Ů" then%>CHECKED<%end if%> name=sex type=radio value=Ů>   
            Ů&nbsp;</SPAN> </td>   
                    <td width="173" height="18"> </td>   
        </tr>   
        <tr>   
          <td width="119" height="18"><font color="#008000"><b>*��������:</b></font></td>   
                    <td width="260" height="18"> 
                      <select size="1" name="Byear">   
            <%for i=1940 to year(date)-3%>   
              <option value="<%=i%>" <%if year(rs_lar("britherday"))=i then%> selected <%end if%>><%=i%></option>   
            <%next%>   
            </select>��<select size="1" name="Bmonth">   
            <%for i=1 to 12%>   
              <option value="<%=i%>" <%if month(rs_lar("britherday"))=i then%> selected <%end if%>><%=i%></option>   
            <%next%>   
            </select>��<select size="1" name="Bday">   
            <%for i=1 to 31%>   
              <option value="<%=i%>" <%if day(rs_lar("britherday"))=i then%> selected <%end if%>><%=i%></option>   
            <%next%>   
            </select>��</td>   
                    <td width="173" height="18"> </td>   
        </tr>   
        <tr>   
          <td width="119" height="11"><font color="#008000"><b>*���ļ���:</b></font></td>   
                    <td width="260" height="11" valign="middle"> 
                      <input type="text" name="home" size="11" style="border-style: solid; border-width: 1" value="<%=rs_lar("home")%>"></td> 
                    <td width="173" height="11" valign="middle"><i>������50���ַ�</i></td> 
        </tr> 
        <tr> 
          <td width="119" height="11"><font color="#008000"><b>*����ѧ��</b></font></td> 
                    <td width="260" height="11"><SPAN class=tt4> 
                      <SELECT name=education 
            size=1> <OPTION selected value=4>����</OPTION> <OPTION 
              value=2>��ר</OPTION> <OPTION value=3>��ר</OPTION> <OPTION 
              value=1>����</OPTION> <OPTION value=6>��ʿ</OPTION> <OPTION 
              value=5>˶ʿ</OPTION> <OPTION value=0>��У��</OPTION> <OPTION 
              value=7>����</OPTION></SELECT> </SPAN></td> 
                    <td width="173" height="11"></td> 
        </tr> 
        <tr> 
          <td width="119" height="19"><font color="#008000"><b>*����ְҵ:</b></font></td> 
                    <td width="260" height="19"> 
                      <SELECT name="job" size="1"> 
              <OPTION selected value=����>����</OPTION> <OPTION value=����>����</OPTION>
              <OPTION value=���ڲ���>���ڲ���</OPTION> <OPTION value=����>����</OPTION> <OPTION
              value=����ҵ��>����ҵ��</OPTION> <OPTION value=����>����</OPTION> <OPTION
              value=����>����</OPTION> <OPTION value=��ʦ>��ʦ</OPTION> <OPTION
              value=����>����</OPTION> <OPTION value=���>���</OPTION> <OPTION
              value=����>����</OPTION> <OPTION value=��������>��������</OPTION> <OPTION
              value=����>����</OPTION> <OPTION value=���ӹ���>���ӹ���</OPTION> <OPTION
              value=IT>IT</OPTION> <OPTION value=����Ա>����Ա</OPTION> <OPTION
              value=����>����</OPTION></SELECT> </td>
                    <td width="173" height="19"> </td>
        </tr>
        <tr>
          <td width="119" height="18"><font color="#008000"><b>&nbsp;</b></font><b>���ĵ�λ:</b></td>
                    <td width="260" height="18"> 
                      <input type="text" name="company" size="31" style="border-style: solid; border-width: 1" value="<%=rs_lar("company")%>"></td>
                    <td width="173" height="18"><i>������50���ַ�</i></td>
        </tr>
        <tr>
          <td width="119" height="18"><font color="#008000"><b>*�����ʱ�:</b></font></td>
                    <td width="260" height="18"> 
                      <input type="text" name="postcalcode" size="6" style="border-style: solid; border-width: 1" value="<%=rs_lar("postcalcode")%>"></td>
                    <td width="173" height="18"><i>6������</i></td>
        </tr>
        <tr>
          <td width="119" height="14"><font color="#008000"><b>*���ĵ绰:</b></font></td>
                    <td width="260" height="14"> 
                      <input type="text" name="tel" size="15" style="border-style: solid; border-width: 1" value="<%=rs_lar("tel")%>"></td>
                    <td width="173" height="14"><i>������20���ַ�</i></td>
        </tr>
        <tr>
          <td width="119" height="11"></td>
                    <td width="260" height="11"></td>
                    <td width="173" height="11"></td>
        </tr>
        <tr>
                    <td height="14" colspan="3"> 
                      <table border="0" width="100%" cellspacing="0" cellpadding="0">
              <tr>
                <td width="100%" style="border-bottom-style: solid" bordercolor="#000000"><b><font color="#FF9900">���ĸ��˼���</font></b></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="119" height="16">&nbsp;���˼���:</td>
                    <td width="260" height="54" rowspan="2"> 
                      <textarea rows="3" name="fresume" cols="37" wrap="hard" style="border-style: solid; border-width: 1"><%=rs_lar("fresume")%></textarea></td>
                    <td width="173" height="19"><i>100������</i></td>
        </tr>
        <tr>
          <td width="119" height="37"></td>
                    <td width="173" height="44"></td>
        </tr>
        <tr>
          <td width="119" height="10"></td>
                    <td width="260" height="10"></td>
                    <td width="173" height="10"></td>
        </tr>
        <tr>
                    <td height="5" colspan="3"> 
                      <table border="0" width="100%" cellspacing="0" cellpadding="0">
              <tr>
                <td width="100%" style="border-bottom-style: solid" bordercolor="#000000"><b><font color="#FF9900">������������</font></b></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="119" height="5"><font color="#008000"><b>*��������:</b></font></td>
                    <td width="260" height="5"> 
                      <input disabled type="text" name="netname" size="11" style="border-style: solid; border-width: 1" value="<%=rs_lar("netname")%>"></td>
                    <td width="173" height="5"><i>2-10�ַ�</i></td>
        </tr>
        <tr>
          <td width="119" height="8"><font color="#008000"><b>&nbsp;</b></font><b>������ҳ:</b></td>
                    <td width="260" height="8"> 
                      <input type="text" name="homepage" size="30" style="border-style: solid; border-width: 1" value="<%=rs_lar("homepage")%>"></td>
                    <td width="173" height="8">Ҫ��д<font color="#FF0000"><b>http://</b></font></td>
        </tr>
        <tr>
          <td width="119" height="5"><font color="#008000"><b>*�����ʼ�:</b></font></td>
                    <td width="260" height="5"> 
                      <input type="text" name="email" size="20" style="border-style: solid; border-width: 1" value="<%=rs_lar("email")%>"></td>
                    <td width="173" height="5"><i>һ��Ҫ��ʵ�����ղ���������Ϣ</i></td>
        </tr>
        <tr>
          <td width="119" height="3"><font color="#008000"><b>&nbsp;</b></font><b>����Ѱ����:</b></td>
                    <td width="260" height="3"> 
                      <input type="text" name="netcall" size="14" style="border-style: solid; border-width: 1" value="<%=rs_lar("netcall")%>">(����Ѱ�������룩</td>
                    <td width="173" height="3"><i>������12���ַ�</i></td>
        </tr>
        <tr>
          <td width="119" height="16"><font color="#008000"><b>&nbsp;</b></font><b>��ȥ��������:</b></td>
                    <td width="260" height="16"> 
                      <input type="text" name="chatroom" size="20" style="border-style: solid; border-width: 1 value=" <%=rs_lar("chatroom")%>""></td>
                    <td width="173" height="16"><i>������50���ַ�</i></td>
        </tr>
        <tr>
          <td width="119" height="11"></td>
                    <td width="260" height="11"></td>
                    <td width="173" height="11"></td>
        </tr>
        <tr>
                    <td height="4" colspan="3"> 
                      <table border="0" width="100%" cellspacing="0" cellpadding="0">
              <tr>
                <td width="100%" style="border-bottom-style: solid" bordercolor="#000000"><b><font color="#FF9900">�����Ը��밮��</font></b></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="119" height="5"><font color="#008000"><b>*ϲ�����˶�:</b></font></td>
                    <td width="260" height="5"> 
                      <input type="text" name="sport" size="20" style="border-style: solid; border-width: 1" value="<%=rs_lar("sport")%>"></td>
                    <td width="173" height="5"><i>������30���ַ�</i></td>
        </tr>
        <tr>
          <td width="119" height="7"><font color="#008000"><b>*ϲ�����鼮:</b></font></td>
                    <td width="260" height="7"> 
                      <input type="text" name="book" size="20" style="border-style: solid; border-width: 1" value="<%=rs_lar("book")%>"></td>
                    <td width="173" height="7"><i>������50���ַ�</i></td>
        </tr>
        <tr>
          <td width="119" height="6"><font color="#008000"><b>*ϲ��������:</b></font></td>
                    <td width="260" height="6"> 
                      <input type="text" name="music" size="25" style="border-style: solid; border-width: 1" value="<%=rs_lar("music")%>"></td>
                    <td width="173" height="6"><i>������50���ַ�</i></td>
        </tr>
        <tr>
          <td width="119" height="5"><font color="#008000"><b>*ϲ��������:</b></font></td>
                    <td width="260" height="5"> 
                      <input type="text" name="people" size="20" style="border-style: solid; border-width: 1" value="<%=rs_lar("people")%>"></td>
                    <td width="173" height="5"><i>������30���ַ�</i></td>
        </tr>
        <tr>
          <td width="119" height="5"><font color="#008000"><b>&nbsp;</b></font><b>�������û��س�:</b></td>
                    <td width="260" height="5"> 
                      <input type="text" name="interest" size="30" style="border-style: solid; border-width: 1" value="<%=rs_lar("interest")%>"></td>
                    <td width="173" height="5"><i>������50���ַ�</i></td>
        </tr>
        <tr>
          <td width="119" height="11"><font color="#008000"><b>*��������:</b></font></td>
                    <td width="260" height="11"> 
                      <input type="text" name="adage" size="31" style="border-style: solid; border-width: 1" value="<%=rs_lar("adage")%>"></td>
                    <td width="173" height="11"><i>������50���ַ�</i></td>
        </tr>
        <tr>
          <td width="119" height="9"><font color="#008000"><b>*�Ը�����:</b></font></td>
                    <td width="260" height="9"> 
                      <input type="text" name="character" size="31" style="border-style: solid; border-width: 1" value="<%=rs_lar("character")%>"></td>
                    <td width="173" height="9"><i>������50���ַ�</i></td>
        </tr>
        <tr>
          <td width="119" height="16">��</td>
                    <td width="260" height="16"> 
                      <p align="right"></td>
                    <td width="173" height="16"> 
                      <p align="right"><input type="submit" value="�ύ" name="B1"><input type="reset" value="��д" name="B2">
          </td>
        </tr>
      </table>
    </td>
</table>
</div>
</form>


    </td>
    <td width="1%" valign="top">

    </td>
    <td width="25%" valign="top"> 
      <table border="1" width="100%" bordercolor="#CC3300" cellspacing="0" cellpadding="0" height="707">
        <tr bgcolor="#067B0F"> 
          <td width="100%" height="13"> 
            <p align="center"><font color="#FFFFFF"><b>�����ѵĻ�</b></font></p>
        </td>
      </tr>
      <tr>
        <td width="100%" height="690" valign="top">
            <table border="0" width="100%" height="162" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="100%" height="241" valign="top">&nbsp;&nbsp; Ϊ����������������ʵ�����ң�����ϸ���������д���ĸ��˵����� 
                  <p>&nbsp;&nbsp; ��������ϴ���Ƭ�����ύ�˸��˵����Ժ󣬽���<a href="sendphoto.asp">[��Ƭ�ϴ�]</a>��Ŀ��</p>
                  <p>&nbsp;&nbsp; ���������ĳ�����2001��4�£������ѵõ��˺ܶ����ѵ�֧�֡��ڴ���������֧��������վ�����ѱ�ʾ��л��</p>
                  <p>&nbsp;&nbsp;&nbsp; ףԸ���е����Ѷ����ҵ�����֪����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                  </p>
                </td>
              </tr>
              <tr> 
                <td width="100%" height="1" valign="top"></td>
              </tr>
            </table>
        </td>
      </tr>
    </table>

    </td>
  </tr>
</table>
<map name="Map"> 
  <area shape="rect" coords="226,94,278,110" href="default.asp">
  <area shape="rect" coords="291,93,360,112" href="your.asp">
  <area shape="rect" coords="381,93,446,112" href="list.asp">
  <area shape="rect" coords="469,92,532,113" href="register.asp">
  <area shape="rect" coords="554,91,618,111" href="sendphoto.asp">
  <area shape="rect" coords="640,92,704,112" href="pics.asp">
</map>
</body>

</html>
<%
rs_lar.close
set rs_lar=nothing
set conn=nothing
%>