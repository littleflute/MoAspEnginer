<!--#include file="conn.asp"-->
<!--#include file="upload.inc"-->
<html>
<head>
<STYLE>TD {
	FONT-SIZE: 9pt; LINE-HEIGHT: 140%
}
BODY {
	FONT-SIZE: 9pt; LINE-HEIGHT: 140%
}
A:link {
	COLOR: #0033cc; TEXT-DECORATION: none
}
A:visited {
	COLOR: #0033cc; TEXT-DECORATION: none
}
A:active {
	COLOR: #ff0000; TEXT-DECORATION: none
}
A:hover {
	COLOR: #000000; TEXT-DECORATION: underline
}
</STYLE>
<title>�����ϴ�</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<table width=100% border=0  cellspacing="0" cellpadding="0"><tr><td>
<script>
parent.document.forms[0].submit.disabled=false;
</script>
<%
Server.ScriptTimeOut=999999
dim upload_type
dim upload,file,formName,filename,fileExt
dim ranNum

sql="select * from config"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
upload_type=rs("upload_type")
upload_size=rs("upload_size")
filetype=rs("upload_filetype")
rs.close
set rs = nothing

if upload_type<>0 and upload_type<>1 and upload_type<>2 and upload_type<>3 then
 upload_type=0
end if

if upload_size<1 then
 upload_size=1
end if

if filetype="" then
 filetype="htw,ida,idq,asp,cdx,asa,htr,idc,shtm,shtml,stm,php,php3,aspx"
end if

select case upload_type
case 0
	call upload_0()
case 1
	call upload_1()
case 2
	call upload_2()
case else
	conn.close
	set conn = nothing
	response.write "����Ա�ѽ��ϴ����ܹر�"
	response.end
end select

'�ж����������Ƿ����Ҫ��
Private Function CheckFileExt (fileEXT)
filetype=split(filetype,",")
	for i=0 to ubound(filetype)
		if lcase(fileEXT)=lcase(trim(filetype(i))) then
			CheckFileExt=false
			exit Function
		else
			CheckFileExt=true
		end if
	next
End Function

'===========������ϴ�(upload_0)====================
sub upload_0()
set upload=new UpFile_Class ''�����ϴ�����
upload.GetDate (upload_size*1024)   'ȡ���ϴ�����
 
if upload.err > 0 then
    select case upload.err
	case 1
	Response.Write "������ѡ����Ҫ�ϴ������ϡ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
	case 2
	Response.Write "����̫����ѹ�������ϴ������ϴ�С���ó���"&upload_size&"K����[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
	end select
	set upload=nothing
	exit sub
else

for each formName in upload.file ''�г������ϴ��˵�����
 set file=upload.file(formName)  ''����һ�����϶���
 fileExt=lcase(file.FileExt)
'�ж���������
 if CheckFileExt(fileEXT)=false then
	set upload = nothing
	set file = nothing
 	response.write "����Ա��ֹ�ϴ����������ϣ��������޸���չ�������ϴ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
	exit sub
 end if

randomize
ranNum=int(90000*rnd)+10000
 filename="files/"&session("teacherid")&"at"&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
 if file.FileSize>0 then         ''��� FileSize > 0 ˵������������
  file.SaveToFile Server.mappath(filename)   ''��������
  response.write "�ϴ��ɹ�!"
  response.write "<script>parent.document.forms[0].fileurl.value='"&FileName&"'</script>"
 else
  Response.Write "�����ϴ���СΪ0�����ϡ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
 end if
 set file=nothing
next
set upload=nothing
end if
end sub

'===========Lyfupload����ϴ�(upload_1)=========================
sub upload_1()
	dim obj,filename,fileExt_a
	dim ss
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'���ϴ�С����
    	obj.maxsize=upload_size*1024
	
	if obj.request("fname")="" or isnull(obj.request("fname")) then
	Response.Write "������ѡ����Ҫ�ϴ������ϡ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
	exit sub
	end if
	randomize
	ranNum=int(90000*rnd)+10000
	fileExt_a=split(obj.request("fname"),".")
	fileExt=lcase(fileExt_a(ubound(fileExt_a)))

	if CheckFileExt(fileEXT)=false then
		set obj = nothing
 		response.write "����Ա��ֹ�ϴ����������ϣ��������޸���չ�������ϴ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		exit sub
	end if

	filename=session("teacherid")&"at"&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
	filename=filename&"."&fileExt

	ss=obj.SaveFile("file1",Server.MapPath("files/"), true,filename)

	if ss= "3" then
   		Response.Write ("�������ظ�!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		set obj=nothing
		exit sub
	elseif ss= "0" then
   		Response.Write ("����̫����ѹ�������ϴ������ϴ�С���ó���"&upload_size&"K����[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		set obj=nothing
		exit sub
	elseif ss = "" then
 		Response.Write ("�ϴ�ʧ��!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		exit sub
	else
		Response.Write "�ϴ��ɹ�!" 
		response.write "<script>parent.document.forms[0].fileurl.value='files/"&filename & "'</script>"
		set obj=nothing
		exit sub
	end if
end sub



''===========================Aspupload3.0����ϴ�============================
sub upload_2()
dim Count
on error resume next
Set Upload = Server.CreateObject("Persits.Upload") 
Upload.OverwriteFiles = false   '����������������
Upload.IgnoreNoPost = True
Upload.SetMaxSize upload_size*1024, True	 '���ϴ�С����

Count = Upload.Save
If Err.Number = 8 Then 
   Response.Write "����̫����ѹ�������ϴ������ϴ�С���ó���"&upload_size&"K����[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]" 
Else 
   If Err <> 0 Then 
      Response.Write "������Ϣ: " & Err.Description 
   Else
		If Count < 1 Then 
		Response.Write "������ѡ����Ҫ�ϴ������ϡ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		set Upload =nothing
		exit sub
		End If
	For Each file in Upload.Files	'�г������ϴ�����
	fileExt=lcase(replace(File.ext,".",""))
	'�ж���������
	if CheckFileExt(fileEXT)=false then
	set upload = nothing
 	response.write "����Ա��ֹ�ϴ����������ϣ��������޸���չ�������ϴ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
	exit sub
	end if
	'���ϱ�����ֵ
	randomize
	ranNum=int(90000*rnd)+10000
	filename=session("teacherid")&"at"&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&File.ext
	filename="files/"&filename
	file.saveas Server.MapPath(filename)	'�ϴ���������
	response.Write "�ϴ��ɹ�!" 
	response.write "<script>parent.document.forms[0].fileurl.value='" &filename& "'</script>"
	Next
   End If 
End If
set Upload = nothing
end sub

conn.close
set conn = nothing
%>
</td></tr></table>
</body>
</html>