<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/md5code.asp" -->
<!--#include file="header.asp" -->
<%
Dim msg_Title,msg_Content
IF Request.QueryString("action")="postcomm" Then
	Dim blog_ID
	blog_ID=Request.Form("blog_ID")
	IF IsInteger(blog_ID)=False Then
		msg_Title="出现错误"
		msg_Content="<a href=""javascript:history.go(-1);"">参数出现错误，点击返回上一页</a>"
	ElseIf (memStatus<>"SupAdmin" And memStatus<>"Admin") And DateDiff("s",Request.Cookies(CookieName)("memLastPost"),Now())<15 Then
		msg_Title="出现错误"
		msg_Content="<a href=""javascript:history.go(-1);"">你发表评论速度太快了，点击返回上一页</a>"
	ElseIf Trim(Session("L-Blog_ValidateCode"))<>Trim(Request.Form("validatecode")) Then
		msg_Title="出现错误"
		msg_Content="<a href=""javascript:history.go(-1);"">请输入发表评论按钮旁边的验证码框，点击返回上一页</a>"
	Else
		Dim comm_LogQuery,comm_LogISOK
		Set comm_LogQuery=Conn.ExeCute("SELECT log_DisComment,log_IsShow FROM blog_Content WHERE log_ID="&blog_ID&"")
		IF comm_LogQuery.EOF AND comm_LogQuery.BOF Then
			comm_LogISOK=1
		Else
			IF comm_LogQuery(0)=True OR comm_LogQuery(1)=False Then
				comm_LogISOK=2
			End IF
		End IF
		Set comm_LogQuery=Nothing
		Dim comm_AllreadyMem,comm_AllreadyMemErr
		Set comm_AllreadyMem=Server.CreateObject("ADODB.RecordSet")
		SQL="SELECT mem_Name,mem_Password,mem_Status,mem_LastIP FROM blog_Member WHERE mem_Name='"&CheckStr(Request.Form("comm_memName"))&"'"
		comm_AllreadyMem.Open SQL,Conn,1,3
		SQLQueryNums=SQLQueryNums+1
		IF comm_AllreadyMem.EOF AND comm_AllreadyMem.BOF Then
			comm_AllreadyMemErr=0
		ElseIF comm_AllreadyMem("mem_Password")=MD5(CheckStr(Request.Form("comm_MemPassword"))) Then
			Response.Cookies(CookieName)("memName")=comm_AllreadyMem("mem_Name")
			Response.Cookies(CookieName)("memPassword")=comm_AllreadyMem("mem_Password")
			Response.Cookies(CookieName)("memStatus")=comm_AllreadyMem("mem_Status")
			memName=comm_AllreadyMem("mem_Name")
			comm_AllreadyMem("mem_LastIP")=Guest_IP
			comm_AllreadyMem.Update
			comm_AllreadyMemErr=2
		Else
			comm_AllreadyMemErr=1
		End IF
		comm_AllreadyMem.Close
		Set comm_AllreadyMem=Nothing
		IF CheckStr(Request.Form("message"))=Empty OR CheckStr(Request.Form("comm_memName"))=Empty Then
			msg_Title="出现错误"
			msg_Content="<a href=""javascript:history.go(-1);"">请将必须信息填写完整，点击返回上一页</a>"
		ElseIF Len(CheckStr(Request.Form("comm_memName")))>24 Then
			msg_Title="出现错误"
			msg_Content="<a href=""javascript:history.go(-1);"">用户名长度超过24个字符，12个汉字，点击返回上一页</a>"
		ElseIF IsValidUserName(CheckStr(Request.Form("comm_memName")))=False Then
			msg_Title="出现错误"
			msg_Content="<a href=""javascript:history.go(-1);"">用户名中含有非法字符，点击返回上一页</a>"
		ElseIF memName=Empty AND comm_AllreadyMemErr=1 Then
			msg_Title="出现错误"
			msg_Content="<a href=""javascript:history.go(-1);"">对不起，你所使用的用户名已经注册，点击返回上一页</a>"
		ElseIF comm_LogISOK=1 Then
			msg_Title="出现错误"
			msg_Content="<a href=""javascript:history.go(-1);"">对不起，你所评论的日志不存在，点击返回上一页</a>"
		ElseIF Not(memStatus="SupAdmin" OR memStatus="Admin") AND comm_LogISOK=2 Then
			msg_Title="出现错误"
			msg_Content="<a href=""javascript:history.go(-1);"">对不起，你所评论的日志不允许发表评论，点击返回上一页</a>"
		Else
			Dim comm_Content,comm_memName,comm_DisSM,comm_DisUBB,comm_DisIMG,comm_AutoURL,comm_AutoKEY
			comm_Content=CheckWordFilter(CheckStr(Request.Form("message")))
			comm_memName=CheckStr(Request.Form("comm_memName"))
			comm_DisSM=Request.Form("comm_DisSM")
			comm_DisUBB=Request.Form("comm_DisUBB")
			comm_DisIMG=Request.Form("comm_DisIMG")
			comm_AutoURL=Request.Form("comm_AutoURL")
			comm_AutoKEY=Request.Form("comm_AutoKEY")
			IF comm_DisSM=Empty Then comm_DisSM=0
			IF comm_DisUBB=Empty Then comm_DisUBB=0
			IF comm_DisIMG=Empty Then comm_DisIMG=0
			IF comm_AutoURL=Empty Then comm_AutoURL=0
			IF comm_AutoKEY=Empty Then comm_AutoKEY=0
			IF memName=Empty And comm_AllreadyMemErr<>2 Then
				Dim comm_SaveMem,comm_MemPassword
				comm_SaveMem=Request.Form("comm_SaveMem")
				comm_MemPassword=MD5(CheckStr(Request.Form("comm_MemPassword")))
				IF comm_SaveMem=1 Then
					Conn.ExeCute("INSERT INTO blog_Member(mem_Name,mem_Password,mem_LastIP) VALUES ('"&comm_memName&"','"&comm_memPassword&"','"&Guest_IP&"')")
					Conn.ExeCute("UPDATE blog_Info SET blog_MemNums=blog_MemNums+1")
					SQLQueryNums=SQLQueryNums+2
					Response.Cookies(CookieName)("memName")=comm_memName
					Response.Cookies(CookieName)("memPassword")=comm_memPassword
					Response.Cookies(CookieName)("memStatus")="Member"
				End IF
				Conn.ExeCute("INSERT INTO blog_Comment(blog_ID,comm_Content,comm_Author,comm_DisSM,comm_DisUBB,comm_DisIMG,comm_AutoURL,comm_AutoKEY,comm_PostIP) VALUES ("&blog_ID&",'"&comm_Content&"','"&comm_Memname&"',"&comm_DisSM&","&comm_DisUBB&","&comm_DisIMG&","&comm_AutoURL&","&comm_AutoKEY&",'"&Guest_IP&"')")
				SQLQueryNums=SQLQueryNums+1
			Else
				Conn.ExeCute("INSERT INTO blog_Comment(blog_ID,comm_Content,comm_Author,comm_DisSM,comm_DisUBB,comm_DisIMG,comm_AutoURL,comm_AutoKEY,comm_PostIP) VALUES ("&blog_ID&",'"&comm_Content&"','"&memName&"',"&comm_DisSM&","&comm_DisUBB&","&comm_DisIMG&","&comm_AutoURL&","&comm_AutoKEY&",'"&Guest_IP&"')")
				SQLQueryNums=SQLQueryNums+1
			End IF
			Application.Lock
			Application.Contents(CookieName&"_blog_LastComm") = ""
			Application.UnLock
			Conn.ExeCute("UPDATE blog_Content SET log_CommNums=log_CommNums+1 WHERE log_ID="&blog_ID&"")
			Conn.ExeCute("UPDATE blog_Member SET mem_PostComms=mem_PostComms+1 WHERE mem_Name='"&comm_memName&"'")
			Conn.ExeCute("UPDATE blog_Info SET blog_CommNums=blog_CommNums+1")
			SQLQueryNums=SQLQueryNums+3
			Response.Cookies(CookieName)("memLastpost")=Now()

			msg_Title="发表成功"
			msg_Content="<a href='blogview.asp?logID="&blog_ID&"'>评论发表成功，点击返回，或者3秒后自动返回</a><meta http-equiv='refresh' content='3;url=blogview.asp?logID="&blog_ID&"'>"
		End IF
	End IF
ElseIF Request.QueryString("action")="delecomm" Then
	IF IsInteger(Request.QueryString("commID"))=False OR IsInteger(Request.QueryString("logID"))=False Then
		msg_Title="出现错误"
		msg_Content="<a href=""javascript:history.go(-1);"">参数出现错误，点击返回上一页</a>"
	Else
		Dim log_AuthorQuery
		Set log_AuthorQuery=Conn.ExeCute("SELECT log_Author FROM blog_Content WHERE log_ID="&CheckStr(Request.QueryString("logID")))
		SQLQueryNums=SQLQueryNums+1
		IF log_AuthorQuery.EOF AND log_AuthorQuery.BOF Then
			msg_Title="出现错误"
			msg_Content="<a href=""javascript:history.go(-1);"">参数出现错误，点击返回上一页</a>"
		Else
			IF Not (memStatus="SupAdmin" OR (memStatus="Admin" And memName=log_AuthorQuery(0))) Then
				msg_Title="出现错误"
				msg_Content="<a href=""javascript:history.go(-1);"">你没有权限删除评论，点击返回上一页</a>"
			Else
				Dim dele_Comm
				Set dele_Comm=Conn.ExeCute("SELECT blog_ID,comm_Author FROM blog_Comment WHERE comm_ID="&CheckStr(Request.QueryString("commID")))
				SQLQueryNums=SQLQueryNums+1
				IF dele_Comm.EOF AND dele_Comm.BOF Then
					msg_Title="出现错误"
					msg_Content="<a href=""javascript:history.go(-1);"">没有找到指定评论，点击返回上一页</a>"
				Else
					Conn.ExeCute("UPDATE blog_Content SET log_CommNums=log_CommNums-1 WHERE log_ID="&dele_Comm("blog_ID"))
					Conn.ExeCute("UPDATE blog_Info SET blog_CommNums=blog_CommNums-1")
					Conn.ExeCute("UPDATE blog_Member SET mem_PostComms=mem_PostComms-1 WHERE mem_Name='"&CheckStr(dele_Comm("comm_Author"))&"'")
					Conn.Execute("DELETE  FROM blog_Comment WHERE comm_ID="&CheckStr(Request.QueryString("commID")))
					SQLQueryNums=SQLQueryNums+4
					Application.Lock
					Application.Contents(CookieName&"_blog_LastComm") = ""
					Application.UnLock
					msg_Title="删除成功"
					msg_Content="<a href='blogview.asp?logID="&CheckStr(Request.QueryString("logID"))&"'>评论删除成功，点击返回</a>"
				End IF
				Set dele_Comm=Nothing
			End IF
		End IF
		Set log_AuthorQuery=Nothing
	End IF
End IF%>
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
    <td width="128" align="center" nowrap><b>发表评论<br>
          <br>
    <font color="#FF0000"><%=msg_Title%></font> </b></td>
    <td align="center" valign="top"><br>
        <br>
      <div class="msg_head"><%=msg_Title%></div>
          <div class="msg_content"><%=msg_Content%></div>
        <br>
    <br></td>
  </tr>
</table>
<!--#include file="footer.asp" -->