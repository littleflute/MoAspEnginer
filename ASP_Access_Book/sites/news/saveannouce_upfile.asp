<!--#include FILE="admin/upload.inc"-->
<html>
<head>
<title>�ļ��ϴ�</title>
</head>
<body   class="tdbg">
<script>

</script>
<%
'�ϴ�������0�������
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
uploadsize=500
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
'�����ϴ�����
set upload=new upload_5xSoft 
 formPath=upload.form("filepath")
 pic_type=cint(upload.form("pic_type"))
 '��Ŀ¼���(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 

response.write "<body class=tdbg leftmargin=5 topmargin=3>"

iCount=0
'�г������ϴ��˵��ļ�
for each formName in upload.file 
'����һ���ļ�����
 set file=upload.file(formName)  
 if file.filesize<100 then
 	response.write "<font size=2>����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if
 	
 if file.filesize>cint(uploadsize)*1000 then
 	response.write "<font size=2>�ļ���С����������100K��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if
'�ж��ļ��ĺ�׺��
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
'��� FileSize > 0 ˵�����ļ�����
 if file.FileSize>0 then         
'�����ļ�
  file.SaveAs Server.mappath(FileName)   
	for i=0 to ubound(Forumupload)
		if fileEXT="."&trim(Forumupload(i)) then
	 	 	response.write "<script>parent.myform.piccontent.value='"&FileName&"'</script>"
	 		if pic_type=0 then
                response.write "<script>parent.myform.txtcontent.value+='[img]"&FileName&"[/img]'</script>"
	 		else
                response.write "<script>parent.myform.txtcontent.value+='<img src="&FileName&" align=right>'</script>"
			end if
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