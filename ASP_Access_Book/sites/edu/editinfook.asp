<!--#include file="conn.asp"-->
<!--#include file="fenlei.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="isteacher.asp"-->
<%
fenlei1 = server.htmlencode(trim(request("fenlei1")))
fenlei2 = server.htmlencode(trim(request("fenlei2")))
teacher = server.htmlencode(trim(request("teacher")))
loginname = trim(request("loginname"))
email = server.htmlencode(trim(request("email")))
homepage = server.htmlencode(trim(request("homepage")))
address = server.htmlencode(trim(request("address")))
intro = server.htmlencode(trim(request("intro")))
photourl = server.htmlencode(trim(request("fileurl")))
qq = server.htmlencode(trim(request("qq")))
password = request("password")
password1 = request("password1")
ask = trim(request("ask"))
answer = trim(request("answer"))

if fenlei1 = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入您所属的"&strfenlei1&"');history.go(-1);</script>"
 response.end
end if
if fenlei2 = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入您所属的"&strfenlei2&"');history.go(-1);</script>"
 response.end
end if
if teacher = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入您的姓名');history.go(-1);</script>"
 response.end
end if
if len(teacher) > 10 then
 conn.close
 set conn = nothing
 response.write "<script>alert('姓名不得超过10个汉字');history.go(-1);</script>"
 response.end
end if
if loginname = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入登录名');history.go(-1);</script>"
 response.end
end if
if len(loginname) > 10 then
 conn.close
 set conn = nothing
 response.write "<script>alert('登录名不得超过10个汉字');history.go(-1);</script>"
 response.end
end if
if IsNumeric(loginname) then
 conn.close
 set conn = nothing
 response.write "<script>alert('登录名不得全为数字');history.go(-1);</script>"
 response.end
end if

set rs = server.createobject("adodb.recordset")
sql = "select count(teacherid) from teacher where loginname='"&loginname&"' and teacherid<>"&session("teacherid")
rs.open sql,conn,1,1
if rs(0) > 0 then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('该登录名已经有人使用了');history.go(-1);</script>"
 response.end
end if
rs.close

if len(email) > 30 then
 conn.close
 set conn = nothing
 response.write "<script>alert('E-Mail地址不得超过30个字符');history.go(-1);</script>"
 response.end
end if
if len(homepage) > 50 then
 conn.close
 set conn = nothing
 response.write "<script>alert('个人主页不得超过50个字符');history.go(-1);</script>"
 response.end
end if
if len(qq) > 10 then
 conn.close
 set conn = nothing
 response.write "<script>alert('QQ号码不应该超过10位吧？');history.go(-1);</script>"
 response.end
end if
if len(address) > 25 then
 conn.close
 set conn = nothing
 response.write "<script>alert('通讯地址不得超过25个汉字');history.go(-1);</script>"
 response.end
end if
if photourl <> "" and left(photourl,6) <> "files" and right(photourl,3) <> "jpg" and right(photourl,3) <> "gif" and right(photourl,3) <> "png" then
 conn.close
 set conn = nothing
 response.write "<script>alert('照片格式不对，只支持jpg、gif和png格式的图片');history.go(-1);</script>"
 response.end
end if
if intro = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入个人简介');history.go(-1);</script>"
 response.end
end if
if password <> password1 then
 conn.close
 set conn = nothing
 response.write "<script>alert('两次密码输入不一致');history.go(-1);</script>"
 response.end
end if
if ask = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入密码找回问题');history.go(-1);</script>"
 response.end
end if
if len(ask) > 25 then
 conn.close
 set conn = nothing
 response.write "<script>alert('密码找回问题不得超过25个汉字');history.go(-1);</script>"
 response.end
end if
if answer = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入密码找回答案');history.go(-1);</script>"
 response.end
end if
if len(answer) > 25 then
 conn.close
 set conn = nothing
 response.write "<script>alert('密码找回答案不得超过25个汉字');history.go(-1);</script>"
 response.end
end if

sql = "select * from teacher where teacherid="&session("teacherid")
rs.open sql,conn,1,3
if rs.bof and rs.eof then
  rs.close
  set rs = nothing
  conn.close
  set conn = nothing
  response.write "<script>alert('请不要捣乱');top.window.location.href='teachermain.asp';</script>"
  response.end
else
  rs("teacher")=teacher
  rs("fenlei1")=fenlei1
  rs("fenlei2")=fenlei2
  rs("email")=email
  if homepage <> "" and left(homepage,7) <> "http://" then
   homepage = "http://"&homepage
  end if
  rs("homepage")=homepage
  rs("address")=address
  rs("photourl")=photourl
  rs("qq")=qq
  rs("intro")=intro
  rs("ask")=ask
  rs("answer")=answer
  rs("loginname")=loginname
  if password <> "" then rs("password")=password
  rs.update
  rs.close
  set rs = nothing
  conn.close
  set conn = nothing
end if
response.write "<script>alert('修改成功');window.location.href='editinfo.asp?id="&session("teacherid")&"';</script>"
%>