<%
	set Conn=Server.CreateObject("ADODB.Connection")

	Conn.open "DBQ="+server.mappath("btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"    
	
	if request("D") <> "" then
		S=request("D")
		SS=split(S,"/")
		SQL = "delete from ��Ʒ��Ϣ where ����='" + SS(0) + "'and ͼ��='" + SS(1) + "'"
		Conn.execute SQL
	end if
	
	if request("C") <> "" then
		SQL="Select * from ��Ʒ��Ϣ Where ����='" + request("C") + "'"
	else
		SQL="Select * from ��Ʒ��Ϣ Where ����='ս�Բ���'"  
	end if
%>



<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>

<body>


<form name="prolist" method="post"><input type="hidden" name="OP" value="�޸�"><div align="center">
<table border="0" width="770" cellspacing="1" cellpadding="0">
<tr>
  <td width="764" colspan="2">
  <font  size="3"><b>ѡ���Ʒ���</b></font>
  </td>
</tr>
<tr>
  <td width="76">
  <font  size="2">�����</font>
  </td>
  <td width="688">
  <font  size="2" color="0000ff">&nbsp;
  	<a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=��ƪ����'">��ƪ����</a>&nbsp; 
  	<a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=��ͳ����'">��ͳ����</a>&nbsp; 
  	<a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=��������'">��������</a>&nbsp; 
  	<a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=��������'">��������</a>
  </font>   
  </td>   
</tr>   
<tr>   
  <td width="76">   
  <font size="2" >�����</font>   
  </td>   
  <td width="688">
   <font size="2"  color="0000ff">&nbsp;
  	<a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=����ѧϰ'">����ѧϰ</a>&nbsp; 
  	<a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=����ѧϰ'">����ѧϰ</a>&nbsp; 
    <a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=��ͥ�ٿ�'">��ͥ�ٿ�</a>&nbsp; 
    <a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=��ý����'">��ý����</a>&nbsp; 
    <a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=�鶾ɱ��'">�鶾ɱ��</a>&nbsp; 
    <a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=�¸�����'">�¸�����</a>
   </font>   
  </td>   
</tr>   
<tr>   
  <td width="76">   
  <font size="2" >��Ϸ��</font>   
  </td>   
  <td width="688">
   <font size="2"  color="0000ff">&nbsp;
  	<a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=��ɫ����'">��ɫ����</a>&nbsp; 
  	<a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=����ð��'">����ð��</a>&nbsp; 
    <a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=ս�Բ���'">ս�Բ���</a>&nbsp; 
    <a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=ģ�⾭Ӫ'">ģ�⾭Ӫ</a>&nbsp; 
    <a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=��������'">��������</a>&nbsp; 
    <a href style="CURSOR: hand" onClick="this.href='pro_list.asp?C=��������'">��������</a>
   </font>   
  </td>   
</tr>   
</table>
<hr>
<table border="0" width="770" cellspacing="1" cellpadding="0">
<tr>   
  <td width="764">   
  </td>   
</tr>   
<tr>   
  <td width="602" align="center" valign="top">   
  <div align="center">   
    <center>   
    <table border="1" cellspacing="0" width="771" bordercolorlight="#FFFFFF" bordercolordark="#FFCCFF" cellpadding="0">   
      <tr>   
        <td width="52" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">ͼ��</font></td>  
        <td width="110" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">����</font></td>  
        <td width="50" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">����</font></td>  
        <td width="54" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">����</font></td>  
        <td width="54" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">��װ</font></td>  
        <td width="30" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">ԭʼ����</font></td>  
        <td width="28" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">�����ۼ�</font></td>  
        <td width="287" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">˵��</font></td>  
        <td width="43" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">���</font></td>  
        <td width="40" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">����</font></td>  
      </tr>  
  
<%  
	set rs=Conn.execute(SQL)  
	Do while not rs.EOF  
		Mstr="prolist.action='pro_list.asp?D=" + rs.fields("����") + "/" + rs.fields("ͼ��") + "'; prolist.submit()"  
		Lstr="sheet.asp?SQL=" + rs.fields("����")
		Lstr=Lstr + "/" + rs.fields("ͼ��")
		Lstr=Lstr + "/" + rs.fields("����")
				
		Lstr="prolist.action='" + Lstr + "';  prolist.submit()"
%>  
      <tr>  
        <td width="52" valign="top" align="left" rowspan="2"><img src="pro-image/<%=rs.fields("ͼ��")%>" width="50" height="60"></td>  
        <td width="110" valign="top" align="left" rowspan="2"><font  size="2"><%=rs.fields("����")%></font></td>  
        <td width="50" valign="top" align="left" rowspan="2"><font  size="2"><%=rs.fields("����")%></font></td>  
        <td width="54" valign="top" align="left" rowspan="2"><font  size="2"><%=rs.fields("����")%></font></td>  
        <td width="54" valign="top" align="left" rowspan="2"><font  size="2"><%=rs.fields("��װ")%></font></td>  
        <td width="30" valign="top" align="left" rowspan="2"><font  size="2"><%=rs.fields("����")%></font></td>  
        <td width="28" valign="top" align="left" rowspan="2"><font  size="2"><%=rs.fields("�ۼ�")%></font></td>  
        <td width="287" valign="top" align="left" rowspan="2"><font  size="2"><%=rs.fields("˵��")%></font></td>  
        <td width="43" valign="top" align="left" rowspan="2"><font  size="2"><%=rs.fields("���")%></font></td>  
        <td width="40" valign="baseline" align="center"><input type="button" value="ɾ��" name="B1" onClick="<%=Mstr%>"></td>  
      </tr>  
      <tr>  
        <td width="42" valign="baseline" align="center"><input type="button" value="�޸�" name="B2" onClick="<%=Lstr%>"></td>  
      </tr>  
  
        
<%  
	rs.movenext  
	loop  
	Conn.close  
%>  
  
  
  
    </table>  
    </center>  
  </div>  
  </td>  
</tr>  
<tr>  
  <td width="764">  
  </td>  
</tr>  
</table></div>  
</form>  
</body>  
  
</html>  
