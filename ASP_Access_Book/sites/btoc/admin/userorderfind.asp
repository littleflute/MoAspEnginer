<%
	Num = Request("T1")
	Nam = Request("T2")
	
	ff = 0 
	SQL = "Select * from �������� where "
	if Num <> "" then 
		SQL = SQL + "������ like '%" + Num + "%'"
		ff = 1 
	end if
	
	if Nam <> "" then
		if ff = 1 then
			SQL = SQL + " and �û��� like '%" + Nam + "%'"
		else	
			SQL = SQL + " �û��� like '%" + Nam + "%'"
			ff = 1
		end if
	end if
	
	

	if ff = 1 then
		set cn=Server.CreateObject("ADODB.Connection")

		cn.open "DBQ="+server.mappath("../btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
		set rs = cn.Execute(SQL)
	end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>���������Ϲ����̳Ƕ�����ѯϵͳ</title>
</head>

<body topmargin="0" leftmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="74%">
    <tr>
      <td width="100%" colspan="5">
        <hr>
      </td>
    </tr>
    <tr>
      <td width="100%" colspan="5">
        <p align="center"><font >������ѯϵͳ</font></td>
    </tr>
    <tr>
      <td width="100%" colspan="5">
        <div align="center"><form name="SS" action="userorderfind.asp">
          <table border="1" cellpadding="0" cellspacing="0" width="80%" bordercolorlight="#FFFFFF" bordercolordark="#808080" height="98">
            <tr>
              <td width="20%" bgcolor="#C0C0C0" align="center" height="21"><font size="2">�ͻ��ʺ�</font></td>
              <td width="38%" height="21"><input type="text" name="T1" size="27"></td>
              <td width="42%"  rowspan="2" valign="top" align="left"><font size="2">&nbsp;<span style="background-color: #FFFF00"><font color="#0000FF">ʹ�÷���:</font></span><br>
                &nbsp; 1.�����ܶ����д�����ݣ��Է���׼ȷ���ҵ��������ݡ�<br>
                &nbsp; 2.���������롱�е���ĸ��ʹ�ô�д��ʽ ��<br>
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                </font></td> 
            </tr> 
            <tr> 
              <td width="20%" bgcolor="#C0C0C0" align="center" height="23"><font size="2">����</font></td> 
              <td width="38%" height="23"><input type="password" name="T2" size="27"></td> 
            </tr> 
            <tr> 
              <td width="20%" bgcolor="#C0C0C0" align="center" height="23"><font size="2">ȷ������</font></td> 
              <td width="38%" height="23"><input type="password" name="T3" size="27"></td> 
            </tr> 

<tr><td colspan="3" align="center">  <input type="button" value="��ʼ��" name="B1" onClick="
                	if (SS.T1.value == '' && SS.T2.value == '' && SS.T3.value == '') 
                	{
                		alert('����д����Ҫ��ѯ������!')
                	} 
                	else
                	{
                		SS.submit()
                	}         
                
                	">&nbsp;&nbsp; <input type="reset" value="����" name="B2">
</td>

<tr>
          </table></form> 
        </div> 
  </center></td> 
    </tr> 
  <center> 
  <tr> 
      <td width="100%" colspan="5"> 
        <p align="center"><b><font face="����_GB2312"><br> 
        </font></b><font >������ѯ���</font></p> 
      </td> 
  </tr> 
  <tr> 
      <td width="100%" colspan="5"> 
        <div align="center"> 
          <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#000080"> 
            <tr> 
              <td width="4%" bgcolor="#FF99FF">&nbsp;</td> 
              <td width="14%" bgcolor="#FF99FF" align="center"><font size="2">��������</font></td> 
              <td width="19%" bgcolor="#FF99FF" align="center"><font size="2">�ͻ��ʺ�</font></td> 
              <td width="12%" bgcolor="#FF99FF" align="center"><font size="2">���ͷ�ʽ</font></td> 
              <td width="12%" bgcolor="#FF99FF" align="center"><font size="2">���ʽ</font></td> 
              <td width="20%" bgcolor="#FF99FF" align="center"><font size="2">��дʱ��</font></td> 
            </tr> 
<% 	
	if ff = 1 then
	Do while Not rs.EOF 
%> 
            <tr> 
              <td width="4%"><font size="2" color="blue"><a href onClick=" 
            	window.open('userformtxt.asp?Num=<%=rs.fields("������")%>&Nam=<%=rs.fields("�û���")%>','formtxt','width=375,height=400,scrollbars=1')" style="CURSOR:hand">.&gt;&gt;</a> 
            	</font></td> 
              <td width="14%"><%=rs.fields("������")%></td> 
              <td width="19%"><%=rs.fields("�û���")%></td> 
              <td width="12%"><%=rs.fields("����")%></td> 
              <td width="12%"><%=rs.fields("����")%></td> 
              <td width="20%"><%=rs.fields("��дʱ��")%></td> 

            </tr> 
<% 
	rs.Movenext 
	Loop 
	cn.Close 
	end if
%> 
          </table> 
        </div> 
    </td> 
  </tr> 
  <tr> 
      <td width="100%" colspan="5"> 
    </td> 
  </tr> 
  </table> 
  </center> 
</div> 
 
</body> 
 
</html> 
