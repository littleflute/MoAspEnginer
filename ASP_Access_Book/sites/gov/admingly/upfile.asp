<%OPTION EXPLICIT%>
<%Server.ScriptTimeOut=5000
Response.Buffer=true%>
<!--#include FILE="upload_5xsoft.inc"-->
<html>
<head>
<title>�ļ��ϴ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body>
<% 

 ''����ǰ�����ں�ʱ��תΪ�ļ���
  Function MakeFileName()
    Dim fname
    fname = Now()
    fname = Trim(fname)
    fname = Replace(fname, "-", "")
    fname = Replace(fname, "/", "")
    fname = Replace(fname, " ", "") 
    fname = Replace(fname, ":", "")
    fname = Replace(fname, "PM", "")
    fname = Replace(fname, "AM", "")
    fname = Replace(fname, "����", "")
    fname = Replace(fname, "����", "")
    MakeFileName = fname
  End Function 

  Dim upload,file,formName,formPath
  Dim i,l,fileType,NewFileName,filenamelist
  '�������ļ�����
  NewFileName = MakeFileName()
  '�����ϴ�����
  Set upload = New upload_5xsoft
  '�ϴ��ļ�Ŀ¼
  formPath = Server.MapPath("movie")&"/"
  '�г������ϴ��˵��ļ�
  For Each formName in upload.objFile
    '����һ���ļ�����
    Set file = upload.file(formName)  
    '��� FileSize > 0 ˵�����ļ�����
    If file.FileSize>10 Then
      'ȡ���ļ���չ��      
      fileType = file.FileName '�ļ����Լ���չ��
      i = instr(fileType,".")  '�Ƿ���ڡ�.��
      l = Len(fileType)  
      If i>0 Then
        fileType = Right(fileType,l-i+1) '�õ���չ��
      End If
      NewFileName = NewFileName & fileType
      filenamelist = formPath & NewFileName  '���ļ����Ե�ַ������
      file.SaveAs filenamelist   ''�����ļ�
    else
	response.write"<script>alert('û�з����ϴ����ļ�');</script>"
	newfilename=""
	End If
    Set file = Nothing
  next
  set upload=nothing  ''ɾ���˶���
Response.Write "<SCRIPT>parent.frmCtoy.filename.value='"&newfilename&"'</SCRIPT>"
Response.Write "<font style='font-family: ����; font-size: 8pt'>�ļ��ϴ��ɹ�,�밴ȷ����Ť<br>"
%>
</body>
</html>