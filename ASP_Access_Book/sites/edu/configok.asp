<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="isadmin.asp"-->
<%
on error resume next
schoolname = trim(request("schoolname"))
sn = trim(request("sn"))
adminpwd = request("adminpwd")
adminpwd1 = request("adminpwd1")
strfenlei1 = trim(request("strfenlei1"))
strfenlei2 = trim(request("strfenlei2"))
adminmail = trim(request("adminmail"))
contactadmin = trim(request("contactadmin"))
siteurl = trim(request("siteurl"))
locked = trim(request("locked"))
pagesize = int(trim(request("pagesize")))
upload_type = int(trim(request("upload_type")))
upload_size = int(trim(request("upload_size")))
upload_filetype = trim(request("upload_filetype"))
headads = request("headads")
gonggao = request("gonggao")
if schoolname = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入学校名称');history.go(-1);</script>"
 response.end
end if
if len(schoolname) > 25 then
 conn.close
 set conn = nothing
 response.write "<script>alert('学校名称不得超过25个汉字');history.go(-1);</script>"
 response.end
end if
if sn = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入授权序列号');history.go(-1);</script>"
 response.end
end if
if len(sn) <> 13 then
 conn.close
 set conn = nothing
 response.write "<script>alert('序列号错误');history.go(-1);</script>"
 response.end
end if
if strfenlei1 = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入一级分类名');history.go(-1);</script>"
 response.end
end if
if len(strfenlei1) > 5 then
 conn.close
 set conn = nothing
 response.write "<script>alert('一级分类名不得超过5个汉字');history.go(-1);</script>"
 response.end
end if
if strfenlei2 = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入二级分类名');history.go(-1);</script>"
 response.end
end if
if len(strfenlei2) > 5 then
 conn.close
 set conn = nothing
 response.write "<script>alert('二级分类名不得超过5个汉字');history.go(-1);</script>"
 response.end
end if
if adminmail = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入管理员E-Mail');history.go(-1);</script>"
 response.end
end if
if len(adminmail) > 100 then
 conn.close
 set conn = nothing
 response.write "<script>alert('管理员E-Mail不得超过100个字符');history.go(-1);</script>"
 response.end
end if
if contactadmin = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入管理员联系方式');history.go(-1);</script>"
 response.end
end if
if len(contactadmin) > 25 then
 conn.close
 set conn = nothing
 response.write "<script>alert('管理员联系方式不得超过25个汉字');history.go(-1);</script>"
 response.end
end if
if siteurl = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入系统URL');history.go(-1);</script>"
 response.end
end if
if left(siteurl,7) <> "http://" then siteurl = "http://"&siteurl
if len(siteurl) > 100 then
 conn.close
 set conn = nothing
 response.write "<script>alert('系统URL不得超过100个字符');history.go(-1);</script>"
 response.end
end if
if adminpwd <> adminpwd1 then
 conn.close
 set conn = nothing
 response.write "<script>alert('两次密码输入不一致');history.go(-1);</script>"
 response.end
end if
if pagesize < 1 then
 response.write "<script>alert('每页显示数目至少为1，系统已自动修正为1')</script>"
 pagesize = 1
end if
if upload_size < 1 then
 response.write "<script>alert('上传大小限制至少为1K，系统已自动修正为1K')</script>"
 upload_size = 1
end if

sql = "select * from config"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs("schoolname")=schoolname
rs("sn")=sn
rs("strfenlei1")=strfenlei1
rs("strfenlei2")=strfenlei2
rs("adminmail")=adminmail
rs("contactadmin")=contactadmin
rs("siteurl")=siteurl
rs("locked")=cint(locked)
rs("pagesize")=pagesize
rs("upload_type")=upload_type
rs("upload_size")=upload_size
rs("upload_filetype")=upload_filetype
rs("headads")=headads
rs("gonggao")=gonggao
if adminpwd <> "" then
 rs("adminpwd")=md5(adminpwd,32)
end if
rs.update
rs.close
set rs = nothing
conn.close
set conn = nothing
response.write "<script>alert('修改成功');window.location.href='config.asp';</script>"
%>