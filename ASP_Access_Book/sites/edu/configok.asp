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
 response.write "<script>alert('������ѧУ����');history.go(-1);</script>"
 response.end
end if
if len(schoolname) > 25 then
 conn.close
 set conn = nothing
 response.write "<script>alert('ѧУ���Ʋ��ó���25������');history.go(-1);</script>"
 response.end
end if
if sn = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('��������Ȩ���к�');history.go(-1);</script>"
 response.end
end if
if len(sn) <> 13 then
 conn.close
 set conn = nothing
 response.write "<script>alert('���кŴ���');history.go(-1);</script>"
 response.end
end if
if strfenlei1 = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('������һ��������');history.go(-1);</script>"
 response.end
end if
if len(strfenlei1) > 5 then
 conn.close
 set conn = nothing
 response.write "<script>alert('һ�����������ó���5������');history.go(-1);</script>"
 response.end
end if
if strfenlei2 = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('���������������');history.go(-1);</script>"
 response.end
end if
if len(strfenlei2) > 5 then
 conn.close
 set conn = nothing
 response.write "<script>alert('�������������ó���5������');history.go(-1);</script>"
 response.end
end if
if adminmail = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('���������ԱE-Mail');history.go(-1);</script>"
 response.end
end if
if len(adminmail) > 100 then
 conn.close
 set conn = nothing
 response.write "<script>alert('����ԱE-Mail���ó���100���ַ�');history.go(-1);</script>"
 response.end
end if
if contactadmin = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('���������Ա��ϵ��ʽ');history.go(-1);</script>"
 response.end
end if
if len(contactadmin) > 25 then
 conn.close
 set conn = nothing
 response.write "<script>alert('����Ա��ϵ��ʽ���ó���25������');history.go(-1);</script>"
 response.end
end if
if siteurl = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('������ϵͳURL');history.go(-1);</script>"
 response.end
end if
if left(siteurl,7) <> "http://" then siteurl = "http://"&siteurl
if len(siteurl) > 100 then
 conn.close
 set conn = nothing
 response.write "<script>alert('ϵͳURL���ó���100���ַ�');history.go(-1);</script>"
 response.end
end if
if adminpwd <> adminpwd1 then
 conn.close
 set conn = nothing
 response.write "<script>alert('�����������벻һ��');history.go(-1);</script>"
 response.end
end if
if pagesize < 1 then
 response.write "<script>alert('ÿҳ��ʾ��Ŀ����Ϊ1��ϵͳ���Զ�����Ϊ1')</script>"
 pagesize = 1
end if
if upload_size < 1 then
 response.write "<script>alert('�ϴ���С��������Ϊ1K��ϵͳ���Զ�����Ϊ1K')</script>"
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
response.write "<script>alert('�޸ĳɹ�');window.location.href='config.asp';</script>"
%>