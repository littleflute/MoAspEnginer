<!--#include file="conn.asp"-->
<%
dim rs_lar
dim sql
dim i
'�Ѷ�Session�����Ƿ�ʱ
if isnull(session("user_id")) then
   response.redirect "timeout.htm"
end if
'�Ѷϴ��û��Ƿ��Ѿ����
if session("user_id")="1" then
    response.redirect "notreg.htm"
	 response.end
end if
'�ж��Ƿ��Ѿ���д����
Set rs_lar = Server.CreateObject("ADODB.Recordset")
sql="select * from larchives where user_id =" & session("user_id")
rs_lar.open sql,conn,3,2
if not(rs_lar.eof and rs_lar.bof) then
   response.redirect "haveregist.htm"
   response.end
end if
rs_lar.close
set rs_lar=nothing
set conn=nothing
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>��������</title>
<style>
<!--
-->
</style>
<link rel="stylesheet" href="ysb.css">
</head>

<body leftmargin="5" topmargin="0">
<table width="772" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="images/top.gif" width="772" height="112" usemap="#Map" border="0"></td>
  </tr>
</table>
<table border="0" width="776" cellspacing="0" cellpadding="0">
  <tr align="center"> 
    <td style="border-bottom-style: solid" width="136" height="19"><a href="default.asp" style="color:#ffffff"><font size="2"><font color="#000000">������ҳ</font></font></a></td>
    <td style="border-bottom-style: solid" width="126" height="19"><a href="reg.asp" class="x">�������</a></td>
    <td style="border-bottom-style: solid" width="126" height="19"><a href="register.asp" class="x">����ע��</a></td>
    <td style="border-bottom-style: solid" width="126" height="19"><a href="your.asp" class="x">����ר��</a></td>
    <td style="border-bottom-style: solid" width="126" height="19"><a href="list.asp" class="x">��������</a></td>
    <td style="border-bottom-style: solid" width="117" height="19"><a href="sendphoto.asp" class="x">��Ƭ�ϴ�</a></td>
    <td style="border-bottom-style: solid" width="9" height="19">&nbsp;</td>
  </tr>
</table>
<table border="0" width="776" cellspacing="0" cellpadding="0" bgcolor="#FFEBE1">
  <tr> 
    <td width="552" valign="top"> 
      <form action="ladd.asp" method="POST">
        <div align="left"> 
          <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr valign="top"> 
              <td width="600"> 
                <table border="0" width="100%" cellspacing="0" cellpadding="0">
                  <tr bgcolor="#067B0F" align="center"> 
                    <td height="19" colspan="3"><font color="#FFFFFF"><b>��������д���ĸ�������</b><font color="#FFCCCC">����*�ŵı�����д�� 
                      </font></font></td>
                </tr>
                <tr> 
                  <td height="4" colspan="3"><font color="#FFFFFF">&nbsp;</font></td>
                </tr>
                <tr> 
                  <td height="7" style="border-bottom-style: solid" bordercolor="#000000" colspan="3"><b><font color="#FF9900">���Ļ�������</font></b></td>
                </tr>
                <tr> 
                  <td width="119" height="18"><font color="#008000"><b>*��ʵ����:</b></font></td>
                  <td width="294" height="18"> 
                    <input type="text" name="name" size="10" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="18"><i>2-10�ַ�</i></td>
                </tr>
                <tr> 
                  <td width="119" height="18"><font color="#008000"><b>*�����Ա�:</b></font></td>
                  <td width="294" height="18"> 
                    <INPUT CHECKED name=sex type=radio
            value=��>
                    <SPAN class=tt4>��&nbsp; 
                    <INPUT name=sex type=radio value=Ů>
                    Ů&nbsp;</SPAN> </td>
                  <td width="139" height="18"> </td>
                </tr>
                <tr> 
                  <td width="119" height="18"><font color="#008000"><b>*��������:</b></font></td>
                  <td width="294" height="18"> 
                    <select size="1" name="Byear">
                      <%for i=1940 to year(date)-3%> 
                      <option value="<%=i%>"><%=i%></option>
                      <%next%> 
                    </select>
                    �� 
                    <select size="1" name="Bmonth">
                      <%for i=1 to 12%> 
                      <option value="<%=i%>"><%=i%></option>
                      <%next%> 
                    </select>
                    �� 
                    <select size="1" name="Bday">
                      <%for i=1 to 31%> 
                      <option value="<%=i%>"><%=i%></option>
                      <%next%> 
                    </select>
                    ��</td>
                  <td width="139" height="18"> </td>
                </tr>
                <tr> 
                  <td width="119" height="11"><font color="#008000"><b>*���ļ���:</b></font></td>
                  <td width="294" height="11" valign="middle"><SPAN class=tt4> 
                    <SELECT class=small13
            name=province size=1>
                      <option value="�Ϻ���">�Ϻ���</option>
                      <option value="������">������</option>
                      <option value="�����">�����</option>
                      <option value="�ӱ�ʡ">�ӱ�ʡ</option>
                      <option value="ɽ��ʡ">ɽ��ʡ</option>
                     
                      <option value="����ʡ">����ʡ</option>
                      <option value="����ʡ">����ʡ</option>
                      <option value="������ʡ">������ʡ</option>
                      <option value="����ʡ">����ʡ</option>
                      <option value="�㽭ʡ">�㽭ʡ</option>
                      <option value="����ʡ">����ʡ</option>
                      <option value="����ʡ">����ʡ</option>
                      <option value="����ʡ">����ʡ</option>
                      <option value="ɽ��ʡ">ɽ��ʡ</option>
                      <option value="����ʡ">����ʡ</option>
                      <option value="����ʡ">����ʡ</option>
                      <option value="����ʡ" >����ʡ</option>
                      
                      <option value="�Ĵ�ʡ">�Ĵ�ʡ</option>
                      <option value="����ʡ">����ʡ</option>
                      <option value="����ʡ">����ʡ</option>
                     
                      <option value="����ʡ">����ʡ</option>
                      <option value="����ʡ">����ʡ</option>
                      <option value="�ຣʡ">�ຣʡ</option>
                     
                      <option value="������">������</option>
                      <option value="����ʡ">����ʡ</option>
                      <option value="����">����</option>
                      <option value="����">����</option>
                      <option value="������">������</option>
                      <option value="�㶫ʡ" selected>�㶫ʡ</option>
                    </SELECT>
                    </SPAN> 
                    <input type="text" name="home" size="11" style="border-style: solid; border-width: 1">
                    (��/����</td>
                  <td width="139" height="11" valign="middle"><i>(��/��)2-10�ַ�</i></td>
                </tr>
                <tr> 
                  <td width="119" height="11"><font color="#008000"><b>*����ѧ��</b></font></td>
                  <td width="294" height="11"><SPAN class=tt4> 
                    <SELECT name=education
            size=1>
                      <OPTION selected value=����>����</OPTION>
                      <OPTION
              value=��ר>��ר</OPTION>
                      <OPTION value=��ר>��ר</OPTION>
                      <OPTION
              value=����>����</OPTION>
                      <OPTION value=��ʿ>��ʿ</OPTION>
                      <OPTION
              value=˶ʿ>˶ʿ</OPTION>
                      <OPTION value=��У��>��У��</OPTION>
                      <OPTION
              value=����>����</OPTION>
                    </SELECT>
                    </SPAN></td>
                  <td width="139" height="11"></td>
                </tr>
                <tr> 
                  <td width="119" height="19"><font color="#008000"><b>*����ְҵ:</b></font></td>
                  <td width="294" height="19"> 
                    <SELECT name="job" size="1">
                      <OPTION selected value=����>����</OPTION>
                      <OPTION value=����>����</OPTION>
                      <OPTION value=���ڲ���>���ڲ���</OPTION>
                      <OPTION value=����>����</OPTION>
                      <OPTION
              value=����ҵ��>����ҵ��</OPTION>
                      <OPTION value=����>����</OPTION>
                      <OPTION
              value=����>����</OPTION>
                      <OPTION value=��ʦ>��ʦ</OPTION>
                      <OPTION
              value=����>����</OPTION>
                      <OPTION value=���>���</OPTION>
                      <OPTION
              value=����>����</OPTION>
                      <OPTION value=��������>��������</OPTION>
                      <OPTION
              value=����>����</OPTION>
                      <OPTION value=���ӹ���>���ӹ���</OPTION>
                      <OPTION
              value=IT>IT</OPTION>
                      <OPTION value=����Ա>����Ա</OPTION>
                      <OPTION
              value=����>����</OPTION>
                    </SELECT>
                  </td>
                  <td width="139" height="19"> </td>
                </tr>
                <tr> 
                  <td width="119" height="18"><font color="#008000"><b>&nbsp;</b></font><b>���ĵ�λ:</b></td>
                  <td width="294" height="18"> 
                    <input type="text" name="company" size="31" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="18"><i>������50���ַ�</i></td>
                </tr>
                <tr> 
                  <td width="119" height="18"><font color="#008000"><b>*�����ʱ�:</b></font></td>
                  <td width="294" height="18"> 
                    <input type="text" name="postcalcode" size="6" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="18"><i>6������</i></td>
                </tr>
                <tr> 
                  <td width="119" height="14"><font color="#008000"><b>*���ĵ绰:</b></font></td>
                    <td width="294" height="14"> 
                      <input type="text" name="tel" size="15" style="border-style: solid; border-width: 1">
                      �������ǲ�����7λ�����֣� </td>
                  <td width="139" height="14"><i>������20���ַ�</i></td>
                </tr>
                <tr> 
                  <td width="119" height="11"></td>
                  <td width="294" height="11"></td>
                  <td width="139" height="11"></td>
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
                  <td width="119" height="20">&nbsp;���˼���:</td>
                  <td width="294" height="1" rowspan="3"> 
                    <textarea rows="4" name="fresume" cols="29" wrap="hard" style="border-style: solid; border-width: 1"></textarea>
                  </td>
                  <td width="139" height="1" rowspan="2"><i>200������</i></td>
                </tr>
                <tr> 
                  <td width="119" height="34"></td>
                </tr>
                <tr> 
                  <td width="119" height="1"></td>
                  <td width="139" height="1"></td>
                </tr>
                <tr> 
                  <td width="119" height="10"></td>
                  <td width="294" height="10"></td>
                  <td width="139" height="10"></td>
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
                  <td width="294" height="5"> 
                    <input type="text" name="netname" size="11" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="5"><i>2-10�ַ�</i></td>
                </tr>
                <tr> 
                  <td width="119" height="8"><font color="#008000"><b>&nbsp;</b></font><b>������ҳ:</b></td>
                  <td width="294" height="8"> 
                    <input type="text" name="homepage" size="30" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="8"><i>������50���ַ�</i></td>
                </tr>
                <tr> 
                  <td width="119" height="5"><font color="#008000"><b>*�����ʼ�:</b></font></td>
                  <td width="294" height="5"> 
                    <input type="text" name="email" size="20" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="5"><i>������50���ַ�</i></td>
                </tr>
                <tr> 
                  <td width="119" height="3"><font color="#008000"><b>&nbsp;</b></font><b>����Ѱ����:</b></td>
                  <td width="294" height="3"> 
                    <select size="1" name="netcall">
                      <option selected value="-û��-">-û��-</option>
                      <option value="ICQ">ICQ</option>
                      <option value="OICQ">OICQ</option>
                    </select>
                    <input type="text" name="netcallcode" size="14" style="border-style: solid; border-width: 1">
                    (����Ѱ�������룩</td>
                  <td width="139" height="3"><i>������12���ַ�</i></td>
                </tr>
                <tr> 
                  <td width="119" height="16"><font color="#008000"><b>&nbsp;</b></font><b>��ȥ��������:</b></td>
                  <td width="294" height="16"> 
                    <input type="text" name="chatroom" size="20" style="border-style: solid; border-width: 1">
                    ����д������ַ��</td>
                  <td width="139" height="16"><i>������50���ַ�</i></td>
                </tr>
                <tr> 
                  <td width="119" height="11"></td>
                  <td width="294" height="11"></td>
                  <td width="139" height="11"></td>
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
                  <td width="294" height="5"> 
                    <input type="text" name="sport" size="20" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="5"><i>������30���ַ�</i></td>
                </tr>
                <tr> 
                  <td width="119" height="7"><font color="#008000"><b>*ϲ�����鼮:</b></font></td>
                  <td width="294" height="7"> 
                    <input type="text" name="book" size="20" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="7"><i>������50���ַ�</i></td>
                </tr>
                <tr> 
                  <td width="119" height="6"><font color="#008000"><b>*ϲ��������:</b></font></td>
                  <td width="294" height="6"> 
                    <input type="text" name="music" size="25" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="6"><i>������50���ַ�</i></td>
                </tr>
                <tr> 
                  <td width="119" height="5"><font color="#008000"><b>*ϲ��������:</b></font></td>
                  <td width="294" height="5"> 
                    <input type="text" name="people" size="20" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="5"><i>������30���ַ�</i></td>
                </tr>
                <tr> 
                  <td width="119" height="5"><font color="#008000"><b>&nbsp;</b></font><b>�������û��س�:</b></td>
                  <td width="294" height="5"> 
                    <input type="text" name="interest" size="30" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="5"><i>������50���ַ�</i></td>
                </tr>
                <tr> 
                  <td width="119" height="11"><font color="#008000"><b>*��������:</b></font></td>
                  <td width="294" height="11"> 
                    <input type="text" name="adage" size="31" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="11"><i>������50���ַ�</i></td>
                </tr>
                <tr> 
                  <td width="119" height="9"><font color="#008000"><b>*�Ը�����:</b></font></td>
                  <td width="294" height="9"> 
                    <input type="text" name="character" size="31" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="9"><i>������50���ַ�</i></td>
                </tr>
                <tr> 
                  <td width="119" height="16">��</td>
                  <td width="294" height="16"> 
                    <p align="right">
                  </td>
                  <td width="139" height="16"> 
                    <p align="right"> 
                      <input type="submit" value="�ύ" name="B1">
                      <input type="reset" value="��д" name="B2">
                  </td>
                </tr>
              </table>
              </td>
            
          </table>
        </div>
      </form>
    </td>
    <td width="10" valign="top"> </td>
    <td width="192" valign="top"> 
      <table border="1" width="100%" bordercolor="#CC0000" cellspacing="0" cellpadding="0" height="707">
        <tr bgcolor="#067B0F"> 
          <td width="100%" height="13"> 
            <p align="center"><font color="#FFFFFF"><b><a href="sendphoto.asp"><font color="#FFCCCC">�ϴ���Ƭ</font></a></b></font></p>
          </td>
        </tr>
        <tr> 
          <td width="100%" height="690" valign="top"> 
            <table border="0" width="100%" height="162" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="100%" height="241" valign="top">&nbsp;&nbsp; <br>
                  Ϊ����������������ʵ�����ң�����ϸ���������д���ĸ��˵����� 
                  <p>&nbsp;&nbsp; ��������ϴ���Ƭ�����ύ�˸��˵����Ժ󣬽���<a href="sendphoto.asp">[��Ƭ�ϴ�]</a>��Ŀ��</p>
                  <p>&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </p>
                </td>
              </tr>
              <tr> 
                <td width="100%" height="17" valign="top"> 
                  <p align="right"> 
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
  <area shape="rect" coords="233,94,273,113" href="default.asp">
  <area shape="rect" coords="293,92,354,110" href="your.asp">
  <area shape="rect" coords="382,93,447,113" href="list.asp">
  <area shape="rect" coords="465,95,531,113" href="register.asp">
  <area shape="rect" coords="553,94,620,113" href="sendphoto.asp">
  <area shape="rect" coords="641,94,704,113" href="pics.asp">
</map>
</body>

</html>
