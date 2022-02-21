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
		fname = replace(fname,"上午","")
	  	fname = replace(fname,"下午","")
	  	fname = int(fname) + int((10-1+1)*Rnd + 1)
		MakedownName=fname
	end function 
formPath="../img/"
iCount=0
for each formName in upload.file ''列出所有上传了的文件
 set file=upload.file(formName)  ''生成一个文件对象
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
newname=MakedownName()&"."&mid(file.FileName,InStrRev(file.FileName, ".")+1)

  file.SaveAs Server.mappath(formPath&newname)   ''保存文件
  iCount=iCount+1
 else 
  response.write "未找到文件 &nbsp;&nbsp;<A HREF=javascript:history.back(1)>返回</A>"
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
<p align="center">文件：</p>
<p align="center">
<%
response.write "<font style=""color:#000000; background-color: #FF0000"">http://"&Request.ServerVariables("SERVER_NAME")&"/img/"&newname&"</font> ("&cint(file.FileSize/1024)&"K) 上传 成功!<br>"
%>
</p>
<p align="center">
请复制红底黑字的文件地址，不要有空格。
</p>
<p align="center">
<a href="upload.asp">继续上传</a>
</p>
<p align="center">
下面是您上传的文件予览
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
set upload=nothing  ''删除此对象
%>
</body>
</html>