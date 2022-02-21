<%OPTION EXPLICIT%>
<%Server.ScriptTimeOut=5000
Response.Buffer=true%>
<!--#include FILE="upload_5xsoft.inc"-->
<html>
<head>
<title>文件上传</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body>
<% 

 ''将当前的日期和时间转为文件名
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
    fname = Replace(fname, "上午", "")
    fname = Replace(fname, "下午", "")
    MakeFileName = fname
  End Function 

  Dim upload,file,formName,formPath
  Dim i,l,fileType,NewFileName,filenamelist
  '创建新文件名称
  NewFileName = MakeFileName()
  '建立上传对象
  Set upload = New upload_5xsoft
  '上传文件目录
  formPath = Server.MapPath("movie")&"/"
  '列出所有上传了的文件
  For Each formName in upload.objFile
    '生成一个文件对象
    Set file = upload.file(formName)  
    '如果 FileSize > 0 说明有文件数据
    If file.FileSize>10 Then
      '取得文件扩展名      
      fileType = file.FileName '文件名以及扩展名
      i = instr(fileType,".")  '是否存在“.”
      l = Len(fileType)  
      If i>0 Then
        fileType = Right(fileType,l-i+1) '得到扩展名
      End If
      NewFileName = NewFileName & fileType
      filenamelist = formPath & NewFileName  '新文件绝对地址和名称
      file.SaveAs filenamelist   ''保存文件
    else
	response.write"<script>alert('没有发现上传的文件');</script>"
	newfilename=""
	End If
    Set file = Nothing
  next
  set upload=nothing  ''删除此对象
Response.Write "<SCRIPT>parent.frmCtoy.filename.value='"&newfilename&"'</SCRIPT>"
Response.Write "<font style='font-family: 宋体; font-size: 8pt'>文件上传成功,请按确定按扭<br>"
%>
</body>
</html>