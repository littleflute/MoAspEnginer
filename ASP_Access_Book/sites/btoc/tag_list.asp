<%
	set Conn=Server.CreateObject("ADODB.Connection")

	Conn.open "DBQ="+server.mappath("btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"    
	
	if request("D") <> "" then
		S=request("D")
		SS=split(S,"/")
		SQL = "Update ��Ʒ��Ϣ set ���='��' where ����='" + SS(0) + "'and ͼ��='" + SS(1) + "'"
		Conn.execute SQL
	end if
	
	if request("C") <> "" then
		S=request("C")
		SS=split(S,"/")
		SQL = "Update ��Ʒ��Ϣ set ���='" + SS(2) + "' where ����='" + SS(0) + "'and ͼ��='" + SS(1) + "'"
		Conn.execute SQL
	end if
%>



<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>

<body>


<form name="taglist" method="post"><div align="center">
<table border="0" width="592" cellspacing="1" cellpadding="0">
<tr>
  <td width="586">
  <font face="����" size="4">
  ��ҳ��Ʒ��ʶ����ǣ�����</font>
  </td>
</tr>
<tr>
  <td width="586">
  <div align="center">
    <center>
    <table border="1" cellspacing="0" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#FFCCFF" cellpadding="0">
      <tr>
        <td width="13%" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">��Ʒͼ��</font></td>
        <td width="39%" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">��Ʒ����</font></td>
        <td width="14%" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">��Ʒ����</font></td>
        <td width="14%" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">��Ʒ��ʶ</font></td>
        <td width="10%" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">����1</font></td>
        <td width="10%" valign="baseline" align="center" bgcolor="#FFCCFF"><font  size="2">����2</font></td>
      </tr>

<%
	SQL="Select * from ��Ʒ��Ϣ Where ���<>'��'"
	set rs=Conn.execute(SQL)
	Do while not rs.EOF
		Mstr="taglist.action='tag_list.asp?D=" + rs.fields("����") + "/" + rs.fields("ͼ��") + "'; taglist.submit()"
		Lstr="var answer=prompt('�������ǣ� ���¡��Ƽ��������','����|�Ƽ�|���'); taglist.action='tag_list.asp?C=" + rs.fields("����") + "/" + rs.fields("ͼ��") + "/' + answer; taglist.submit()"                             
%>
      <tr>
        <td width="13%" valign="baseline" align="center"><img src="pro-image/<%=rs.fields("ͼ��")%>" width="50" height="60"></td>
        <td width="39%" valign="baseline" align="left"><font  size="2"><%=rs.fields("����")%></font></td>
        <td width="14%" valign="baseline" align="center"><font  size="2"><%=rs.fields("����")%></font></td>
        <td width="14%" valign="baseline" align="center"><font  size="2"><%=rs.fields("���")%></font></td>
        <td width="10%" valign="baseline" align="center"><input type="button" value="ɾ��" name="B1" onClick="<%=Mstr%>"></td>
        <td width="10%" valign="baseline" align="center"><input type="button" value="�޸�" name="B2" onClick="<%=Lstr%>"></td>
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
  <td width="586">
  <p align="right"><input type="button" value="�������˵�" name="B3" onClick="location.replace('upld.asp')">
  </td>
</tr>
</table></div>
</form>
</body>

</html>

