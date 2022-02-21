<!--#include FILE="upload.inc"-->
<html>
<head>
<title>文件上传</title>
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
Server.ScriptTimeOut=999999999 '要是你的论坛支持上传的文件比较大，就必须设置。
'上传方式upload_type值: 0＝无组件，1＝lyfupload，2＝Aspupload3.0，3＝chinaaspupload
dim upload_type
upload_type=0

'创建生成预览图片,需要CreatePreviewImage组件支持,upload_view值: 0=不支持,1=支持(根目录下要有PreviewImage文件夹存放文件)
dim upload_view
upload_view=0

'定义变量
dim Forumupload,ranNum
dim formName,formPath,filename,file_name,fileExt,Filesize,F_Type
dim upNum,dateupnum
dim rename,DownloadID

dim previewpath,F_Viewname
F_Viewname=""
previewpath="PreviewImage/"

if  session("AdminUID")="" then
 	response.write "只有登陆用户方能上传！"
	response.end
end if

On Error Resume Next 
select case upload_type
case 0
	call upload_0()
case else
	response.write "本系统未开放插件功能"
	response.end
end select

'===========================无组件上传============================
sub upload_0()
dim upload,file
set upload=new UpFile_Class ''建立上传对象
upload.GetDate (2048*1024)   '取得上传数据,不限大小

if upload.err > 0 then
    select case upload.err
	case 1
	Response.Write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	case 2
	Response.Write "文件大小超过了限制 "&2048&"K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	end select
	exit sub
	else
	formPath=upload.form("filepath")
	'在目录后加(/)
	if right(formPath,1)<>"/" then formPath=formPath&"/"

for each formName in upload.file ''列出所有上传了的文件
	set file=upload.file(formName)  ''生成一个文件对象

	fileExt=lcase(file.FileExt)
	
	'判断文件类型
	'if lcase(fileEXT)="asp" and lcase(fileEXT)="asa" and lcase(fileEXT)="aspx" then
	'	CheckFileExt(fileEXT)=false
	'end if
	'if CheckFileExt(fileEXT)=false then
 '	response.write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	'response.end
	'end if
	if lcase(fileEXT)="asp" or lcase(fileEXT)="asa" or lcase(fileEXT)="aspx" or lcase(fileEXT)=""  then
	response.write "文件格式不正确或者没有选择文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	response.end
	end if	
	'付值变量
	randomize
	ranNum=int(90000*rnd)+10000
	F_Type=CheckFiletype(fileEXT)
	file_name=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
	filename=file_name&"."&fileExt
	rename=filename&"|"
	filename=formPath&filename
	Filesize=file.FileSize

	'记录文件
	if Filesize>0 then         '如果 FileSize > 0 说明有文件数据
	file.SaveToFile Server.mappath(FileName)   ''执行上传文件
'	call checksave()			'记录文件
	end if
	set file=nothing
next
end if
set upload=nothing
'if upNum < int(GroupSetting(40)) and dateupnum < clng(GroupSetting(50)) then
	response.write "文件上传成功 [ <a href=# onclick=history.go(-1)>继续上传</a> ]"
'	else
'	response.write upNum&"个文件上传成功!本次已达到上传数上限。"
'end if
 	response.write "<script>parent.myform.body.value+='[img]"&filename&"[/img]'</script>"
end sub

Private sub checksave()
	'插入上传表并获得ID
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

'判断文件类型是否合格
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

'判断文件类型:0=其它,1=图片,2=FLASH,3=音乐,4=电影
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