<!--#include FILE="upload.inc"-->
<html>
<head>
<title>�ļ��ϴ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
a:visited{text-transform: none; text-decoration: none; font-style: normal; font-weight: normal; font-variant: normal; color: #000000}
a:hover{text-decoration:underline; color: #FF3333}
a:link{text-transform: none; text-decoration: none; font-weight: normal; font-variant: normal; color: #000000}
td { font-size: 9pt; line-height: 13pt; text-decoration: none}
-->
</style>
</head>
<body topmargin="0" bgcolor="#e3f1d1">
<script>
parent.document.forms[0].Submit.disabled=false;
parent.document.forms[0].Submit2.disabled=false;
</script>
<table width="100%" border=0 cellspacing=0 cellpadding=0>
<tr><td class=tablebody2 valign=top height=40>
<%
Server.ScriptTimeOut=999999999 'Ҫ�������̳֧���ϴ����ļ��Ƚϴ󣬾ͱ������á�
'�ϴ���ʽupload_typeֵ: 0���������1��lyfupload��2��Aspupload3.0��3��chinaaspupload
dim upload_type
upload_type=0

'��������Ԥ��ͼƬ,��ҪCreatePreviewImage���֧��,upload_viewֵ: 0=��֧��,1=֧��(��Ŀ¼��Ҫ��PreviewImage�ļ��д���ļ�)
dim upload_view
upload_view=0

'�������
dim Forumupload,ranNum
dim formName,formPath,filename,file_name,fileExt,Filesize,F_Type
dim upNum,dateupnum
dim rename,DownloadID

dim previewpath,F_Viewname
F_Viewname=""
previewpath="PreviewImage/"

if  session("AdminUID")="" then
 	response.write "ֻ�е�½�û������ϴ���"
	response.end
end if

On Error Resume Next 
select case upload_type
case 0
	call upload_0()
case else
	response.write "��ϵͳδ���Ų������"
	response.end
end select

'===========================������ϴ�============================
sub upload_0()
dim upload,file
set upload=new UpFile_Class ''�����ϴ�����
upload.GetDate (2048*1024)   'ȡ���ϴ�����,���޴�С

if upload.err > 0 then
    select case upload.err
	case 1
	Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
	case 2
	Response.Write "�ļ���С���������� "&2048&"K��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
	end select
	exit sub
	else
	formPath=upload.form("filepath")
	'��Ŀ¼���(/)
	if right(formPath,1)<>"/" then formPath=formPath&"/"

for each formName in upload.file ''�г������ϴ��˵��ļ�
	set file=upload.file(formName)  ''����һ���ļ�����

	fileExt=lcase(file.FileExt)
	
	'�ж��ļ�����
	'if lcase(fileEXT)="asp" and lcase(fileEXT)="asa" and lcase(fileEXT)="aspx" then
	'	CheckFileExt(fileEXT)=false
	'end if
	'if CheckFileExt(fileEXT)=false then
 '	response.write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
	'response.end
	'end if
	if lcase(fileEXT)="asp" or lcase(fileEXT)="asa" or lcase(fileEXT)="aspx" or lcase(fileEXT)=""  then
	response.write "�ļ���ʽ����ȷ����û��ѡ���ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
	response.end
	end if	
	'��ֵ����
	randomize
	ranNum=int(90000*rnd)+10000
	F_Type=CheckFiletype(fileEXT)
	file_name=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
	filename=file_name&"."&fileExt
	rename=filename&"|"
	filename=formPath&filename
	Filesize=file.FileSize

	'��¼�ļ�
	if Filesize>0 then         '��� FileSize > 0 ˵�����ļ�����
	file.SaveToFile Server.mappath(FileName)   ''ִ���ϴ��ļ�
'	call checksave()			'��¼�ļ�
	end if
	set file=nothing
next
end if
set upload=nothing
'if upNum < int(GroupSetting(40)) and dateupnum < clng(GroupSetting(50)) then
	response.write "�ļ��ϴ��ɹ� [ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
'	else
'	response.write upNum&"���ļ��ϴ��ɹ�!�����Ѵﵽ�ϴ������ޡ�"
'end if
 	response.write "<script>parent.myform.body.value+='[img]"&filename&"[/img]'</script>"
end sub

Private sub checksave()
	'�����ϴ������ID
	if upload_view=1 and F_Type=1 then
	conn.execute("insert into dv_upfile (F_BoardID,F_UserID,F_Username,F_Filename,F_Viewname,F_FileType,F_Type,F_FileSize) values ("&BoardID&","&UserID&",'"&membername&"','"&replace(rename,"|","")&"','"&F_Viewname&"','"&replace(fileExt,".","")&"',"&F_Type&","&Filesize&")")
	else
	conn.execute("insert into dv_upfile (F_BoardID,F_UserID,F_Username,F_Filename,F_FileType,F_Type,F_FileSize) values ("&BoardID&","&UserID&",'"&membername&"','"&replace(rename,"|","")&"','"&replace(fileExt,".","")&"',"&F_Type&","&Filesize&")")
	end if

	set rs=conn.execute("select top 1 F_ID from dv_upfile order by F_ID desc")
	DownloadID=rs(0)
	rename=DownloadID & ","
	set rs=nothing
	if F_Type=1 then
 	response.write "<script>parent.frmAnnounce.Content.value+='[upload="&fileExt&"]"&filename&"[/upload]'</script>"
	else
 	response.write "<script>parent.frmAnnounce.Content.value+='[upload="&fileExt&"]viewfile.asp?ID="&DownloadID&"[/upload]'</script>"
	end if
	response.write "<script>parent.frmAnnounce.upfilerename.value+='"&rename&"'</script>"

	upNum=upNum+1
	response.cookies("upNum")=upNum
	dateupnum=dateupnum+1
	response.Cookies("dateupnum").Expires=Date+1
	response.cookies("dateupnum")=dateupnum
end sub

'�ж��ļ������Ƿ�ϸ�
Private Function CheckFileExt (fileEXT)
dim Forumupload
Forumupload=split(Board_Setting(19),"|")
	for i=0 to ubound(Forumupload)
		if lcase(fileEXT)=lcase(trim(Forumupload(i))) then
			CheckFileExt=true
			exit Function
		else
			CheckFileExt=false
		end if
	next
End Function

'�ж��ļ�����:0=����,1=ͼƬ,2=FLASH,3=����,4=��Ӱ
Private Function CheckFiletype(fileEXT)
dim upFiletype
dim FilePic,FileVedio,FileSoft,FileFlash,FileMusic
fileEXT=lcase(replace(fileExt,".",""))
FilePic=".gif.jpg.jpeg.png.bmp.tif.iff"
upFiletype=split(FilePic,".")
	for i=0 to ubound(upFiletype)
		if fileEXT=lcase(trim(upFiletype(i))) then
			CheckFiletype=1
			exit Function
		end if
	next
FileFlash=".swf.swi"
upFiletype=split(FileFlash,".")
	for i=0 to ubound(upFiletype)
		if fileEXT=lcase(trim(upFiletype(i))) then
			CheckFiletype=2
			exit Function
		end if
	next
FileMusic=".mid.wav.mp3.rmi.cda"
upFiletype=split(FileMusic,".")
	for i=0 to ubound(upFiletype)
		if fileEXT=lcase(trim(upFiletype(i))) then
			CheckFiletype=3
			exit Function
		end if
	next
FileVedio=".avi.mpg.mpeg.ra.ram.wov.asf"
upFiletype=split(FileVedio,".")
	for i=0 to ubound(upFiletype)
		if fileEXT=lcase(trim(upFiletype(i))) then
			CheckFiletype=4
			exit Function
		end if
	next
FileSoft=".rar.zip.exe.php.php3.asp.aspx.htm.html.shtml.js.jsp.pdf.inc.doc.txt.chm.hlp"
CheckFiletype=0
end function
%>
</td></tr>
</table>
</body>
</html>