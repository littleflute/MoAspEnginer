<!--#include FILE="upload.inc"-->
<%
dim upload,file,formName,formPath,iCount
set upload=new upload_F
	function MakedownName()
		dim fname
	  	fname = now()
		fname = replace(fname,"-","")
	 	fname = replace(fname," ","") 
		fname = replace(fname,":","")
	  	fname = replace(fname,"PM","")
	  	fname = replace(fname,"AM","")
		fname = replace(fname,"����","")
	  	fname = replace(fname,"����","")
	  	fname = int(fname) + int((10-1+1)*Rnd + 1)
		MakedownName=fname
	end function 
formPath="../img/"
iCount=0
for each formName in upload.file ''�г������ϴ��˵��ļ�
 set file=upload.file(formName)  ''����һ���ļ�����
 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
newname=MakedownName()&"."&mid(file.FileName,InStrRev(file.FileName, ".")+1)

  file.SaveAs Server.mappath(formPath&newname)   ''�����ļ�
  iCount=iCount+1
 else 
  response.write "δ�ҵ��ļ� &nbsp;&nbsp;<A HREF=javascript:history.back(1)>����</A>"
  response.end
 end if
next
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../style.css" type="text/css">
</head>

<body>
<p align="center">�ļ���</p>
<p align="center">
<%
response.write "<font style=""color:#000000; background-color: #FF0000"">http://"&Request.ServerVariables("SERVER_NAME")&"/img/"&newname&"</font> ("&cint(file.FileSize/1024)&"K) �ϴ� �ɹ�!<br>"
%>
</p>
<p align="center">
�븴�ƺ�׺��ֵ��ļ���ַ����Ҫ�пո�
</p>
<p align="center">
<a href="upload.asp">�����ϴ�</a>
</p>
<p align="center">
���������ϴ����ļ�����
</p>
<p align="center">
<%
if (mid(file.FileName,InStrRev(file.FileName, ".")+1))="swf" then
%>
<EMBED src="../img/<%=newname%>" quality=high WIDTH=300 HEIGHT=280 TYPE='application/x-shockwave-flash'></EMBED>
<%else%>
<img src="../img/<%=newname%>">
<%end if%>
</p>
<%
set file=nothing
set upload=nothing  ''ɾ���˶���
%>
</body>
</html>