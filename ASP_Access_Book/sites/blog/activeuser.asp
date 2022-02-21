<%
'############### 在线统计脚本说明 ##############

'		此VBScript脚本 专为 LoveYuKi Blog v 1.08 编写 , 需要使用
'Commond.asp文件里的一些变量及常量 , 如果你想在其它地方使用
'而又不知道怎么改 , 请按以下地址联系我 .

'		此脚本利用ASP的Application对像存贮数据 , 不用修改你的数
'据库.

'联系方式 :
'		BLOG : http://www.evolve.name
'		OICQ : 7900132

'		这些信息不会在你的HTML源码中显示出来 ,可以任意转载 , 可以
'删除这些信息 , 我这个小程序 , 版权没有 , 嘿嘿
'		如果你对源代码作了优化工作 , 请通知我 , 谢谢
'2005.3.18.

'Copyright © 2005 Aeon's blog. All Rights Reserved.

'############### ##############  ##########
Const ActiveTime = 15

Call ActiveUserStat()

Call displayUser()

Function ActiveUserStat()
	Dim ActiveUser
	Dim ActiveIP
	Dim ActiveStatus
	Dim memberList,temMemberList
	Dim guestList,temGuestList
	Dim temMember,temChildMem,temGuest,temChildGue
	Dim temCnt
	Dim FindUser
	
	'查找用户的标志变量
	FindUser = False
	
	'从Application对像里读取 会员游客列表
	memberList = Application(CookieName & "_blog_memberList")
	guestList = Application(CookieName & "_blog_guestList")
	'如果会员或游客列表为空则变量重置为空
	If IsEmpty(memberList) Or IsNull(memberList) Then memberList=""
	If IsEmpty(guestList) Or IsNull(guestList) Then guestList=""
	'如果会员或游客列表末尾为$$$,则删除$$$
	If Right(memberList,3) = "$$$" Then	memberList = Left(memberList,Len(memberList)-3)
	If Right(GuestList,3) = "$$$" Then GuestList = Left(guestList,Len(guestList)-3)
	
	'如果读取Cookies用户名为空,刚用户未登录,用户名设为Guest
	If IsNull(memName) Or IsEmpty(memName) Or memName = "" Then
		ActiveUser = "Guest"
		ActiveStatus = "Guest"
	Else
		ActiveUser = memName
		ActiveStatus = memStatus
	End If
	'用户IP,从Commond.asp里的Guest_IP函数读取
	ActiveIP = Guest_IP
	

	If ActiveUser<>"Guest" Then	'如果用户不是游客
		
		If memberList <> "" Then	'如果memberList里有数据
			temMemberList = ""
			temMember = Split(memberList,"$$$")	'把原数据折分为每个用户的数据
			For temCnt = 0 To UBound(temMember)
				temChildMem = Split(temMember(temCnt) ,"$")	'将当前用户数据字符串折分开
				If temChildMem(0) = ActiveUser Then	'如果存贮用户名的字符=用户名
					temChildMem(1) = ActiveIP
					temChildMem(2) = ActiveStatus
					temChildMem(3) = Now()	'更新在线时间为现在的时间
					FindUser = True	'FindUser置为TRUE
				Else
					If DateDiff("n",CDate(temChildMem(3)),Now()) > ActiveTime Then	'如果存贮的时间到现在时间的时差超过规定时间
						'将这个数据段的所有数据换成XXXX
						temChildMem(0) = "XXXX"
						temChildMem(1) = "XXXX"
						temChildMem(2) = "XXXX"
						temChildMem(3) = "XXXX"
					End If
				End if
				
				'将所有的用户资料重组
				temMemberList = temMemberList & temChildMem(0) & "$" & temChildMem(1) & "$" & temChildMem(2) & "$" & temChildMem(3) & "$$$"
				
			Next
			
			If Not FindUser Then
				'如果没有找到用户,添加用户
				temMemberList = temMemberList & ActiveUser & "$" & ActiveIP & "$" & ActiveStatus & "$" & Now() & "$$$"
				'Response.Write "<div align=""center"" style=""color:#FFFFFF"">member finduser("&temChildMem(0)&"): "& FindUser &"  Add User</div>"
			End If
			'删除做了XXXX标记的用户
			temMemberList = Replace(temMemberList,"XXXX$XXXX$XXXX$XXXX$$$","")
		Else
			'如果memberList里没有数据,写入当前用户为第一个用户
			temMemberList = ActiveUser & "$" & ActiveIP & "$" & ActiveStatus & "$" & Now() & "$$$"
		End if
		memberList = temMemberList
		'更新Application
		Application.Lock()
		Application(CookieName & "_blog_memberList") = memberList
		Application.UnLock()
		
		If GuestList<>"" Then
			temGuestList = ""
			temGuest = Split(GuestList,"$$$")
			For temCnt = 0 To UBound(temGuest)
				temChildGue= Split(temGuest(temCnt) ,"$")
				If DateDiff("n",CDate(temChildGue(1)),Now()) > ActiveTime Then
					temChildGue(0) = "XXXX"
					temChildGue(1) = "XXXX"
				End If
				temGuestList = temGuestList & temChildGue(0) & "$" & temChildGue(1) & "$$$"
			Next
			GuestList = Replace(temGuestList,"XXXX$XXXX$$$","")
			Application.Lock()
			Application(CookieName & "_blog_guestList") = guestList
			Application.UnLock()
		End if
	Else
		If GuestList<>"" Then
			temGuestList = ""
			temGuest = Split(GuestList,"$$$")
			For temCnt = 0 To UBound(temGuest)
				temChildGue= Split(temGuest(temCnt) ,"$")
				If temChildGue(0) = ActiveIP Then
					temChildGue(1) = Now
					FindUser = True
				Else
					If DateDiff("n",CDate(temChildGue(1)),Now) > ActiveTime Then
						temChildGue(0) = "XXXX"
						temChildGue(1) = "XXXX"
					End If
				End If
				temGuestList = temGuestList & temChildGue(0) & "$" & temChildGue(1) & "$$$"
			Next
			
			temGuestList = Replace(temGuestList,"XXXX$XXXX$$$","")
			'Response.Write "<div align=""center"" style=""color:#FFFFFF"">Guest finduser: " & FindUser & "  Only Update</DIV>"			
			If Not FindUser Then
				temGuestList = temGuestList & ActiveIP & "$" & Now() & "$$$"
				'Response.Write "<div align=""center"" style=""color:#FFFFFF"">Guest " & ActiveIP & " : " & FindUser & "  Add IP</DIV>"			
			End if
		Else
			temGuestList = ActiveIP & "$" & Now() & "$$$"
		End if
		
		GuestList = temGuestList
		
		Application.Lock()
		Application(CookieName & "_blog_GuestList") = GuestList
		Application.UnLock()
		
		If memberList <> "" Then
			temMemberList = ""
			temMember = Split(memberList,"$$$")
			For temCnt = 0 To UBound(temMember)
				temChildMem = Split(temMember(temCnt) ,"$")
				If DateDiff("n",CDate(temChildMem(3)),Now()) > ActiveTime Then
					temChildMem(0) = "XXXX"
					temChildMem(1) = "XXXX"
					temChildMem(2) = "XXXX"
					temChildMem(3) = "XXXX"
				End if
				temMemberList = temMemberList & temChildMem(0) & "$" & temChildMem(1) & "$" & temChildMem(2) & "$" & temChildMem(3) & "$$$"
			Next
			memberList = Replace(temMemberList,"XXXX$XXXX$XXXX$XXXX$$$","")
			Application.Lock()
			Application(CookieName & "_blog_memberList") = memberList
			Application.UnLock()
		End if
		
	End If	
End Function

Function displayUser()
	Dim temStr
	Dim temCnt
	Dim UserList
	Dim UserCount
	Dim DataList,DataChild
	Dim online_Pic
	Dim memCount,guestCount
	temStr = ""
	memCount = 0
	guestCount = 0
	
	UserList = Application(CookieName & "_blog_memberList")
	If IsNull(UserList) Or IsEmpty(UserList) Then UserList = ""
	If Right(UserList,3) = "$$$" Then UserList = Left(UserList,Len(UserList)-3)
	If UserList <> "" Then
		DataList = Split(UserList,"$$$")
		For temCnt = 0 To UBound(DataList)
			memCount = memCount + 1
			DataChild = Split(DataList(temCnt),"$")
			
			Select Case DataChild(2)
				Case "SupAdmin"
					online_Pic = "<img src=""images/online_Super.gif"" border=""0"">"
				Case "Admin"
					online_Pic = "<img src=""images/online_Admin.gif"" border=""0"">"
				Case Else
					online_Pic = "<img src=""images/online_Member.gif"" border=""0"">"
			End Select
			temStr = temStr & "<span style=""Width:100px;Height:25px;"">"&online_Pic&"&nbsp;<a href=""member.asp?action=view&memName="&Server.URLEncode(DataChild(0))&""" title=""最近活动时间 "&DateDiff("n",CDate(DataChild(3)),Now)&" 分钟前 ("&DataChild(1)&")"" target=""_blank"">" & DataChild(0) & "</a></span>"
		Next
	End If
	
	UserList = CStr(Application(CookieName & "_blog_GuestList"))
	If IsNull(UserList) Or IsEmpty(UserList) Then UserList = ""
	If Right(UserList,3) = "$$$" Then UserList = Left(UserList,Len(UserList)-3)
	If UserList <> "" Then
		DataList = Split(UserList,"$$$")
		For temCnt = 0 To UBound(DataList)
			guestCount = guestCount + 1
			DataChild = Split(DataList(temCnt),"$")
			
			temStr = temStr & "<span style=""Width:100px;Height:25px;""><img src=""images/online_Guest.gif"" border=""0"">&nbsp;<a href=""ipview.asp?ipdata="&DataChild(0)&""" title=""最近活动时间 "&DateDiff("n",CDate(DataChild(1)),Now)&" 分钟前 ("&DataChild(0)&")"" target=""_blank"">Guest</a></span>"
		Next
	End If
	
	Application.Lock()
	Application(CookieName & "_blog_displayOnline") = temStr
	Application(CookieName & "_blog_displayOnlineTitle") = "最近 " & ActiveTime & " 分钟内活动用户 " &Int(memCount+guestCount) & " 人. 其中注册用户 " & memCount & "人. 游客 " & guestCount & " 人."
	Application(CookieName & "_blog_displayOnlineTitle2") =""&Int(memCount+guestCount)&" 人"
	Application.UnLock()
	
End Function

Function DelMember(Byval Member)
	Dim memberList
	Dim temMemberList
	Dim temMember
	Dim temCHildMem
	Dim temCnt
	
	memberList = Application(CookieName & "_blog_memberList")
	If IsEmpty(memberList) Or IsNull(memberList) Then memberList=""
	If Right(memberList,3) = "$$$" Then	memberList = Left(memberList,Len(memberList)-3)
	
	If memberList <> "" Then
		temMemberList = ""
		temMember = Split(memberList,"$$$")
		For temCnt = 0 To UBound(temMember)
			temChildMem = Split(temMember(temCnt) ,"$")
			If LCase(temChildMem(0)) = LCase(Member) Then
				temChildMem(0) = "XXXX"
				temChildMem(1) = "XXXX"
				temChildMem(2) = "XXXX"
				temChildMem(3) = "XXXX"
			End if
			temMemberList = temMemberList & temChildMem(0) & "$" & temChildMem(1) & "$" & temChildMem(2) & "$" & temChildMem(3) & "$$$"
		Next
		temMemberList = Replace(temMemberList,"XXXX$XXXX$XXXX$XXXX$$$","")
		memberList = temMemberList
		'更新Application
		Application.Lock()
		Application(CookieName & "_blog_memberList") = memberList
		Application.UnLock()
	End If
End Function

Function DelGuest(Byval UserIP)
	Dim guestList
	Dim temGuestList
	Dim temGuest
	Dim temChildGue
	Dim temCnt
	
	guestList = Application(CookieName & "_blog_guestList")
	If IsEmpty(guestList) Or IsNull(guestList) Then guestList=""
	If Right(guestList,3) = "$$$" Then	guestList = Left(guestList,Len(guestList)-3)
	
	If guestList <> "" Then
		temGuestList = ""
		temGuest = Split(guestList,"$$$")
		For temCnt = 0 To UBound(temGuest)
			temChildGue = Split(temGuest(temCnt) ,"$")
			If temChildGue(0) = UserIP Then
				temChildGue(0) = "XXXX"
				temChildGue(1) = "XXXX"
			End if
			temGuestList = temGuestList & temChildGue(0) & "$" & temChildGue(1) & "$$$"
		Next
		temGuestList = Replace(temGuestList,"XXXX$XXXX$$$","")
		guestList = temGuestList
	
		'更新Application
		Application.Lock()
		Application(CookieName & "_blog_guestList") = guestList
		Application.UnLock()
	End If
End Function
%>
