<!--#include FILE="admin/upload.inc"-->
<html>
<head>
<title>�ļ��ϴ�</title>
</head>
<body   class="tdbg">
<script>
//parent.document.forms[0].Submit.disabled=false;
//parent.document.forms[0].Submit2.disabled=false;
</script>
<%
'�ϴ�������0���������1��lyfupload��2��aspupload��3��chinaaspupload
dim upload_type
upload_type=0

dim upNum
dim uploadsuc
dim Forumupload
dim ranNum
dim uploadfiletype
dim upload,file,formName,formPath,iCount,filename,fileExt,pic_type

'�ϴ��ļ���С100K
dim uploadsize
uploadsize=100

upNum=session("upNum")
'�����ϴ�����
'if upNum>3 then
if FALSE then
 	response.write "<font size=2>һ��ֻ���ϴ������ļ���</font>"
	response.end
end if
response.write "<FONT color=red>"&upNum&"</font>"
select case upload_type
case 0
	call upload_0()
case 1
	call upload_1()
case else
	response.write "��ϵͳδ���Ų������"
	response.end
end select

sub upload_0()
set upload=new upload_5xSoft ''�����ϴ�����

 formPath=upload.form("filepath")
 pic_type=cint(upload.form("pic_type"))
 ''��Ŀ¼���(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 

response.write "<body class=tdbg leftmargin=5 topmargin=3>"

iCount=0
for each formName in upload.file ''�г������ϴ��˵��ļ�
 set file=upload.file(formName)  ''����һ���ļ�����
 if file.filesize<100 then
 	response.write "<font size=2>����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if
 	
 if file.filesize>cint(uploadsize)*1000 then
 	response.write "<font size=2>�ļ���С����������100K��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))
 uploadsuc=false
 Forumupload=split("gif,jpg,doc",",")
 for i=0 to ubound(Forumupload)
	if fileEXT="."&trim(Forumupload(i)) then
	uploadsuc=true
	exit for
	else
	uploadsuc=false
	end if
 next
 if uploadsuc=false then
 	response.write "<font size=2>�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if

 randomize
 ranNum=int(90000*rnd)+10000
 filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&fileExt

 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
  file.SaveAs Server.mappath(FileName)   ''�����ļ�
	for i=0 to ubound(Forumupload)
		if fileEXT="."&trim(Forumupload(i)) then
	 	'response.write "<script>parent.myform.piccontent.value+='[upload="&Forumupload(i)&"]"&FileName&"[/upload]'</script>"
	 	response.write "<script>parent.myform.piccontent.value='"&FileName&"'</script>"
	 	'	if pic_type=0 then
                'response.write "<script>parent.myform.txtcontent.value+='[img]"&FileName&"[/img]'</script>"
	 	'	else
                'response.write "<script>parent.myform.txtcontent.value+='<img src="&FileName&" align=right>'</script>"
		'	end if
                exit for
		end if
	next
 iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing

Htmend iCount&" ���ļ��ϴ�����!"
end sub

sub HtmEnd(Msg)
 if upNum="" then upNum=1
 session("upNum")=upNum+1
 response.write "<font size=2>�ļ��ϴ��ɹ� [ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
 response.end
end sub

'lyfupload
sub upload_1()
	dim obj,uploadpath,filename
	dim ss,aa1,aa
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'��С
    	obj.maxsize=Cint(uploadsize)*1000
	'����
    	obj.extname=Forum_upload
	'������
	uploadpath=Server.MapPath(obj.request("filepath"))
	'response.write obj.extname
	'response.end
	filename=MakedownName(obj.filetype("file1"))
	'response.write filename
	ss=obj.SaveFile("file1",uploadpath, true,filename)

	if ss= "3" then
   		Response.Write ("<font size=2>�ļ����ظ�!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		response.end
	elseif ss= "0" then
   		Response.Write ("<font size=2>�ļ��ߴ����!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		response.end
	elseif ss = "1" then
 		Response.Write ("<font size=2>�ļ�����ָ�������ļ�!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		response.end
	elseif ss = "" then
 		Response.Write ("<font size=2>�ļ��ϴ�ʧ��!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		response.end
	else
		if upNum="" then upNum=1
		session("upNum")=upNum+1
		response.write "<font size=2>�ļ��ϴ��ɹ� [ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	 	response.write "<script>parent.myform.piccontent.value+='[upload="&uploadfiletype&"]"&obj.request("filepath") & "/" & filename&"[/upload]'</script>"
		'response.write "<script>parent.document.forms[0].myface.value='" & obj.request("filepath") & "/" & filename & "'</script>"
		session("upface")="done"
		response.end
	end if
	set obj=nothing
end sub

'--------������ת�����ļ���--------
function MakedownName(filename)
  dim fname,Forumupload,u
  fname = now()
  fname = replace(fname,"-","")
  fname = replace(fname," ","") 
  fname = replace(fname,":","")
  fname = replace(fname,"PM","")
  fname = replace(fname,"AM","")
  fname = replace(fname,"����","")
  fname = replace(fname,"����","")
  fname = int(fname) + int((10-1+1)*Rnd + 1)
  Forumupload=split(Forum_upload,",")
  for u=0 to ubound(Forumupload)
		if instr(filename,Forumupload(u))>0 then
			uploadfiletype=Forumupload(u)
			MakedownName=fname & "." & Forumupload(u)
			exit for
		end if
  next
end function
%>
</body>
</html>