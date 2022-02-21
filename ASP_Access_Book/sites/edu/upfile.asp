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
<title>资料上传</title>
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
	response.write "管理员已将上传功能关闭"
	response.end
end select

'判断资料类型是否符合要求
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

'===========无组件上传(upload_0)====================
sub upload_0()
set upload=new UpFile_Class ''建立上传对象
upload.GetDate (upload_size*1024)   '取得上传数据
 
if upload.err > 0 then
    select case upload.err
	case 1
	Response.Write "请首先选择您要上传的资料　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	case 2
	Response.Write "资料太大，请压缩后再上传（资料大小不得超过"&upload_size&"K）　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	end select
	set upload=nothing
	exit sub
else

for each formName in upload.file ''列出所有上传了的资料
 set file=upload.file(formName)  ''生成一个资料对象
 fileExt=lcase(file.FileExt)
'判断资料类型
 if CheckFileExt(fileEXT)=false then
	set upload = nothing
	set file = nothing
 	response.write "管理员禁止上传该类型资料，请打包或修改扩展名后再上传　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	exit sub
 end if

randomize
ranNum=int(90000*rnd)+10000
 filename="files/"&session("teacherid")&"at"&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有资料数据
  file.SaveToFile Server.mappath(filename)   ''保存资料
  response.write "上传成功!"
  response.write "<script>parent.document.forms[0].fileurl.value='"&FileName&"'</script>"
 else
  Response.Write "不得上传大小为0的资料　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
 end if
 set file=nothing
next
set upload=nothing
end if
end sub

'===========Lyfupload组件上传(upload_1)=========================
sub upload_1()
	dim obj,filename,fileExt_a
	dim ss
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'资料大小限制
    	obj.maxsize=upload_size*1024
	
	if obj.request("fname")="" or isnull(obj.request("fname")) then
	Response.Write "请首先选择您要上传的资料　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	exit sub
	end if
	randomize
	ranNum=int(90000*rnd)+10000
	fileExt_a=split(obj.request("fname"),".")
	fileExt=lcase(fileExt_a(ubound(fileExt_a)))

	if CheckFileExt(fileEXT)=false then
		set obj = nothing
 		response.write "管理员禁止上传该类型资料，请打包或修改扩展名后再上传　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		exit sub
	end if

	filename=session("teacherid")&"at"&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
	filename=filename&"."&fileExt

	ss=obj.SaveFile("file1",Server.MapPath("files/"), true,filename)

	if ss= "3" then
   		Response.Write ("资料名重复!　[ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		set obj=nothing
		exit sub
	elseif ss= "0" then
   		Response.Write ("资料太大，请压缩后再上传（资料大小不得超过"&upload_size&"K）　[ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		set obj=nothing
		exit sub
	elseif ss = "" then
 		Response.Write ("上传失败!　[ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		exit sub
	else
		Response.Write "上传成功!" 
		response.write "<script>parent.document.forms[0].fileurl.value='files/"&filename & "'</script>"
		set obj=nothing
		exit sub
	end if
end sub



''===========================Aspupload3.0组件上传============================
sub upload_2()
dim Count
on error resume next
Set Upload = Server.CreateObject("Persits.Upload") 
Upload.OverwriteFiles = false   '不允许覆盖重名资料
Upload.IgnoreNoPost = True
Upload.SetMaxSize upload_size*1024, True	 '资料大小限制

Count = Upload.Save
If Err.Number = 8 Then 
   Response.Write "资料太大，请压缩后再上传（资料大小不得超过"&upload_size&"K）　[ <a href=# onclick=history.go(-1)>重新上传</a> ]" 
Else 
   If Err <> 0 Then 
      Response.Write "错误信息: " & Err.Description 
   Else
		If Count < 1 Then 
		Response.Write "请首先选择你要上传的资料　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		set Upload =nothing
		exit sub
		End If
	For Each file in Upload.Files	'列出所有上传资料
	fileExt=lcase(replace(File.ext,".",""))
	'判断资料类型
	if CheckFileExt(fileEXT)=false then
	set upload = nothing
 	response.write "管理员禁止上传该类型资料，请打包或修改扩展名后再上传　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	exit sub
	end if
	'资料变量付值
	randomize
	ranNum=int(90000*rnd)+10000
	filename=session("teacherid")&"at"&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&File.ext
	filename="files/"&filename
	file.saveas Server.MapPath(filename)	'上传保存资料
	response.Write "上传成功!" 
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