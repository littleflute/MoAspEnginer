<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>

<body>
<div align="center"><form name="S" method="POST"><input type="hidden" name="OP" value="<%=request("OP")%>">

<%
	if request("OP") = "�޸�" then
		set Conn=Server.CreateObject("ADODB.Connection")

		Conn.open "DBQ="+server.mappath("btoc.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"  
		
		response.write "�޸ļ�¼"
		S=request("SQL")
		T=split(S,"/")
		SQL="Select * from ��Ʒ��Ϣ where ����='" + T(0) + "' and ͼ��='" + T(1) + "' and ����='" + T(2) + "'"		
		set rs=Conn.execute(SQL)
		
		TP=rs.fields("ͼ��")
		T1=rs.fields("����")
		T2=rs.fields("����")
		T3=rs.fields("��װ")
		T4=rs.fields("����")
		T5=rs.fields("�ۼ�")
		T6=rs.fields("����")
		T7=rs.fields("���")
		T8=rs.fields("˵��")
		T9=rs.fields("ʹ��")
		C=rs.fields("���")
		Conn.close
	else
		response.write "<font color=""blue""><u>�½���¼</u></font>"
		TP="picture.jpg"
		C="����"
	end if

%>

<table border="1" width="450" bordercolorlight="#FFFFFF" bordercolordark="#FFFF00">
  <tr>
    <td width="338"><font  size="2">����: <input type="text" name="T1" size="41" value="<%=T1%>"></font></td>                                                              
    <td width="96" rowspan="4" align="right"><img src="pro-image/<%=TP%>" width="75" height="90">     
     
    </td>                                                              
  </tr>                                                              
  <tr>                                                              
    <td width="338"><font  size="2">����: <input type="text" name="T2" size="41" value="<%=T2%>"></font></td>                                                              
  </tr>                                                              
  <tr>                                                              
    <td width="338"><font  size="2">��װ: <input type="text" name="T3" size="17" value="<%=T3%>">&nbsp;&nbsp;&nbsp;&nbsp;    
      ����:                                                               
      <select size="1" name="D">                                                              
                      <option selected value="<%=T4%>"><%=T4%></option>                                                    
                      <option value="��ƪ����/C">��ƪ����</option>
                      <option value="��ͳ����/C">��ͳ����</option>
                      <option value="��������/C">��������</option>
                      <option value="��������/C">��������</option>
                      <option value="����ѧϰ/S">����ѧϰ</option>
                      <option value="����ѧϰ/S">����ѧϰ</option>
                      <option value="��ͥ�ٿ�/S">��ͥ�ٿ�</option>
                      <option value="��ý����/S">��ý����</option>
                      <option value="�鶾ɱ��/S">�鶾ɱ��</option>
                      <option value="�¸�����/S">�¸�����</option>
                      <option value="��ɫ����/G">��ɫ����</option>
                      <option value="����ð��/G">����ð��</option>
                      <option value="ս�Բ���/G">ս�Բ���</option>
                      <option value="ģ�⾭Ӫ/G">ģ�⾭Ӫ</option>
                      <option value="��������/G">��������</option>
                      <option value="��������/G">��������</option>
                    </select></font></td>
  </tr>
  <tr>
    <td width="338"><font  size="2">�����ۼ�: <input type="text" name="T5" size="17" value="<%=T5%>">&nbsp;&nbsp;<br>
      ԭʼ����:                                                               
      <input type="text" name="T6" size="11" value="<%=T6%>"></font></td>                                                              
  </tr>                                                              
  <tr>                                                            
    <td width="440" valign="top" align="left" colspan="2"><font  size="2">���</font><font  size="2">:</font></td>                                                             
  </tr>                                                           
  <tr>                                                           
    <td width="440" valign="top" align="left" colspan="2"><textarea rows="4" name="S1" cols="60"><%=T7%></textarea></td>                                                             
  </tr>                                                           
  <tr>                                                             
    <td width="440" valign="top" align="left" colspan="2"><font  size="2">��ϸ˵��:</font></td>                                                             
  </tr>                                                           
  <tr>                                                             
    <td width="440" valign="top" align="left" colspan="2"><textarea rows="4" name="S2" cols="60"><%=T8%></textarea></td>                                                             
  </tr>                                                           
  <tr>                                                             
    <td width="440" valign="top" align="left" colspan="2"><font  size="2">ʹ��˵����</font></td>                                                             
  </tr>                                                           
  <tr>                                                             
    <td width="440" valign="top" align="left" colspan="2"><textarea rows="4" name="S3" cols="60"><%=T9%></textarea></td>                                                             
  </tr>                                                           
  <tr>                                                             
    <td width="338" valign="top" align="left"><font  size="2">���:                                                               
      <input type="text" name="T7" size="18" value="<%=C%>"></font></td>                                                              
    <td width="96" valign="top" align="right"><input type="button" value="ȷ��" name="B1" onClick="                                
			if (S.T1.value=='' || S.T2.value=='' || S.T3.value=='' || S.D.value=='' || S.T5.value=='' || S.T6.value=='' || S.T7.value=='' || S.S1.value=='' || S.S2.value=='' || S.S3.value=='')                                 
    			{                                 
    				alert('����д���еĿհף�')                                 
   				}                               
   			else                         
   				{                         
   					S.action='done.asp'                         
   					S.submit()                         
   				}">                                                          
      <input type="reset" value="����" name="B2"></td>                                                            
  </tr>                                                            
</table>                                                          
                                                          
</form>                                                            
</div>                                                            
</body>                                                            
                                                            
</html>                                                            
