<%
dim conn   
dim connstr
connstr="DBQ="+server.mappath("guest.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
conn.open connstr
'�޸�:
'--------------------------------------------------------------------
home="../"  '�����ҳ,���Ų���ȥ��
Mypage =10  '����ÿҳ��ʾ��������
PassWord = "111"  '��������
%>

