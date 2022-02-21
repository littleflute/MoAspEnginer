<!--#include FILE="admin/upload.inc"-->
<html>
<head>
<title>文件上传</title>
</head>
<body   class="tdbg">
<script>

</script>
<%
'上传方法，0＝无组件
dim upload_type
upload_type=0

dim upNum
dim uploadsuc
dim Forumupload
dim ranNum
dim uploadfiletype
dim upload,file,formName,formPath,iCount,filename,fileExt,pic_type
'上传文件大小100K
dim uploadsize
uploadsize=500
upNum=session("upNum")
'屏蔽上传数量
'if upNum>3 then
if FALSE then
 	response.write "<font size=2>一次只能上传三个文件！</font>"
	response.end
end if
response.write "<FONT color=red>"&upNum&"</font>"
select case upload_type
case 0
	call upload_0()
case 1
	call upload_1()
case else
	response.write "本系统未开放插件功能"
	response.end
end select
sub upload_0()
'建立上传对象
set upload=new upload_5xSoft 
 formPath=upload.form("filepath")
 pic_type=cint(upload.form("pic_type"))
 '在目录后加(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 

response.write "<body class=tdbg leftmargin=5 topmargin=3>"

iCount=0
'列出所有上传了的文件
for each formName in upload.file 
'生成一个文件对象
 set file=upload.file(formName)  
 if file.filesize<100 then
 	response.write "<font size=2>请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if
 	
 if file.filesize>cint(uploadsize)*1000 then
 	response.write "<font size=2>文件大小超过了限制100K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if
'判断文件的后缀名
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
 	response.write "<font size=2>文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if
 randomize
 ranNum=int(90000*rnd)+10000
 filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&fileExt
'如果 FileSize > 0 说明有文件数据
 if file.FileSize>0 then         
'保存文件
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
Htmend iCount&" 个文件上传结束!"
end sub
sub HtmEnd(Msg)
 if upNum="" then upNum=1
 session("upNum")=upNum+1
 response.write "<font size=2>文件上传成功 [ <a href=# onclick=history.go(-1)>继续上传</a> ]</font>"
 response.end
end sub
'--------将日期转化成文件名--------
function MakedownName(filename)
  dim fname,Forumupload,u
  fname = now()
  fname = replace(fname,"-","")
  fname = replace(fname," ","") 
  fname = replace(fname,":","")
  fname = replace(fname,"PM","")
  fname = replace(fname,"AM","")
  fname = replace(fname,"上午","")
  fname = replace(fname,"下午","")
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