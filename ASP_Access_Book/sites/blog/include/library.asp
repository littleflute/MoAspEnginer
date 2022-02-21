<%Sub Calendar(C_Year,C_Month,C_Day)  'BLOG日历
	ReDim Link_Days(2,0)
	Dim Link_Count
	Link_Count=0
	Dim This_Year,This_Month,This_Day,RS_Month,Link_TF
	IF C_Year=Empty Then C_Year=Year(Now())
	IF C_Month=Empty Then C_Month=Month(Now())
	IF C_Day=Empty Then C_Day=0
	C_Year=Cint(C_Year)
	C_Month=Cint(C_Month)
	C_Day=Cint(C_Day)
	This_Year=C_Year
	This_Month=C_Month
	This_Day=C_Day
	Dim To_Day,To_Month,To_Year
	To_Day=Cint(Day(Now()))
	To_Month=Cint(Month(Now()))
	To_Year=Cint(Year(Now()))
	
	SQL="SELECT log_PostYear,log_PostMonth,log_PostDay FROM blog_Content WHERE log_PostYear="&C_Year&" AND log_PostMonth="&C_Month&" ORDER BY log_PostDay"
	Set RS_Month=Server.CreateObject("ADODB.RecordSet")
	RS_Month.Open SQL,Conn,1,1
	SQLQueryNums=SQLQueryNums+1
	Dim the_Day
	the_Day=0
	Do While NOT RS_Month.EOF
		IF RS_Month("log_PostDay")<>the_Day Then
			the_Day=RS_Month("log_PostDay")
			ReDim PreServe Link_Days(2,Link_Count)
			Link_Days(0,Link_Count)=RS_Month("log_PostMonth")
			Link_Days(1,Link_Count)=RS_Month("log_PostDay")
			Link_Days(2,Link_Count)="default.asp?log_Year="&RS_Month("log_PostYear")&"&log_Month="&RS_Month("log_PostMonth")&"&log_Day="&RS_Month("log_PostDay")
			Link_Count=Link_Count+1
		End IF
		RS_Month.MoveNext
	Loop
	RS_Month.Close
	Set RS_Month=Nothing
	
	Dim Month_Name(12)
	Month_Name(0)=""
	Month_Name(1)="1"
	Month_Name(2)="2"
	Month_Name(3)="3"
	Month_Name(4)="4"
	Month_Name(5)="5"
	Month_Name(6)="6"
	Month_Name(7)="7"
	Month_Name(8)="8"
	Month_Name(9)="9"
	Month_Name(10)="10"
	Month_Name(11)="11"
	Month_Name(12)="12"
	
	Dim Month_Days(12)
	Month_Days(0)=""
	Month_Days(1)=31
	Month_Days(2)=28
	Month_Days(3)=31
	Month_Days(4)=30
	Month_Days(5)=31
	Month_Days(6)=30
	Month_Days(7)=31
	Month_Days(8)=31
	Month_Days(9)=30
	Month_Days(10)=31
	Month_Days(11)=30
	Month_Days(12)=31
	
	If IsDate("February 29, " & This_Year) Then Month_Days(2)=29
	
	Dim Start_Week
	Start_Week=WeekDay(C_Month&"-1-"&C_Year)-1
	
	Dim Next_Month,Next_Year,Pro_Month,Pro_Year
	Next_Month=C_Month+1
	Next_Year=C_Year
	IF Next_Month>12 then 
		Next_Month=1
		Next_Year=Next_Year+1
	End IF
	Pro_Month=C_Month-1
	Pro_Year=C_Year
	IF Pro_Month<1 then 
		Pro_Month=12
		Pro_Year=Pro_Year-1
	End IF
	
	Response.Write("<div class=""siderbar_head""><img src=""images/sider_calendar.gif"" align=""absmiddle"" border=""0"" /> 站点日历</div><table width=""100%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"" background=""images/calendar/month"&Month_Name(C_Month)&".gif""><tr><td colspan=""7"" align=""center""><a href=""default.asp?log_Year="&C_Year-1&""" title=""上一年""><span class=""arrow"">7</span></a><a href=""default.asp?log_Year="&Pro_Year&"&log_Month="&Pro_Month&""" title=""上一月""><span class=""arrow"">3</span></a> <strong>"&C_Year&"  - "&Month_Name(C_Month)&"</strong> <a href=""default.asp?log_Year="&Next_Year&"&log_Month="&Next_Month&""" title=""下一月""><span class=""arrow"">4</span></a><a href=""default.asp?log_Year="&C_Year+1&""" title=""下一年""><span class=""arrow"">8</span></a></td></tr><tr bgcolor=""#C8C7BD"" class=""calendar-week"">")
	Response.Write("<td>日</td><td>一</td><td>二</td><td>三</td><td>四</td><td>五</td><td>六</td></tr><tr>")
	Dim i,j,k,l,m
	For  i=0 TO Start_Week-1
		Response.Write("<td>&nbsp;</td>")
	Next
	Dim This_BGColor
	j=1
	While j<=month_Days(This_Month)
	 	For k=start_Week To 6
			This_BGColor="calendar"
			IF j=To_Day AND This_Year=To_Year AND This_Month=To_Month Then This_BGColor="calendar-today"
			IF j=This_Day Then This_BGColor="calendar-thisday"
			Response.Write("<td class="""&This_BGColor&""">")
			Link_TF="Flase"
			For l=0 TO Ubound(Link_Days,2)
				IF Link_Days(0,l)<>"" Then
					IF Link_Days(0,l)=This_Month AND Link_Days(1,l)=j Then
						Response.Write("<a href="""&Link_Days(2,l)&""">")
						Link_TF="True"
					End IF
				End IF
			Next
		IF j<=Month_Days(This_Month) Then Response.Write(j)
		IF Link_TF="True" Then Response.Write("</a>")
        Response.Write("</td>")
		j=j+1
	Next
	Start_Week=0
	Response.Write("</tr>")
	Wend
	Response.Write("</table>")
End Sub

Sub MemberCenter '用户中心
	IF memName=Empty Then
		Response.Write("<div class=""siderbar_head""><img src=""images/sider_member.gif"" align=""absmiddle"" border=""0"" /> 用户登陆</div><div class=""siderbar_main""><center><form name=""memLogin"" action=""logging.asp?action=login"" method=""post"">用户名：<input name=""username"" type=""text"" id=""username"" value="""" size=""12"" maxlength=""20"" /><br>密&nbsp;&nbsp;&nbsp;码：<input name=""Password"" type=""password"" id=""Password"" value="""" size=""12"" maxlength=""20"" /><br>验证码：<input name=""validatecode"" type=""text"" id=""validatecode"" class=""txt"" size=""3"" />&nbsp;<img src=""include/validatecode.asp"" align=""absmiddle"" border=""0"" /><br><input name=""Login"" type=""submit"" id=""Login"" value="" 登 陆 "" />&nbsp;<input name=""Regedit"" type=""button"" id=""Regedit"" value="" 注 册 "" onclick=""javascript:document.location.href='register.asp';"" /></form></center></div>")
	Else
      Response.Write("<div class=""siderbar_head""><img src=""images/sider_member.gif"" align=""absmiddle"" border=""0"" /> 用户中心</div><div class=""siderbar_main"">你好，"&memName&"<br>")
	  IF memStatus="SupAdmin" Then
		Response.Write("<a href=""blogpost.asp""><img src=""images/icon_newblog.gif"" align=""absmiddle"" border=""0"" /> 发表日志</a>&nbsp;|&nbsp;<a href=""admincp.asp"" target=""_blank""><img src=""images/icon_admincp.gif"" align=""absmiddle"" border=""0"" /> 系统管理</a><br><img src=""images/icon_admincp.gif"" align=""absmiddle"" border=""0"" /> <a href=""photoadd.asp"">添加图片</a><br>")
	  ElseIF memStatus="Admin" Then
	  	Response.Write("<a href=""blogpost.asp""><img src=""images/icon_newblog.gif""align=""absmiddle"" border=""0"" /> 发表日志</a><br>")
	  End IF
	  Response.Write("<a href=""favorite.asp""><img src=""images/icon_favorite.gif""align=""absmiddle"" border=""0"" /> 网络书签</a>&nbsp;|&nbsp;<br /><a href=""member.asp?action=edit""><img src=""images/icon_memedit.gif""align=""absmiddle"" border=""0"" /> 修改资料</a>&nbsp;|&nbsp;<a href=""logging.asp?action=logout""><img src=""images/icon_logout.gif"" align=""absmiddle"" border=""0"" /> 退出登录</a></div>")
	End IF
End Sub

Sub NewCommList '最新留言列表
	Response.Write("<div class=""siderbar_head""><img src=""images/sider_newcomm.gif"" border=""0"" align=""absmiddle"" /> 最新评论</div><div class=""siderbar_main"">")
	Dim Arr_LastComm '最新评论缓存
	IF Not IsArray(Application(CookieName&"_blog_LastComm")) Then
		Dim log_LastComm,log_LastCommList
		Set log_LastCommList=Server.CreateObject("ADODB.RecordSet")
		SQL="SELECT TOP 10 C.comm_ID,C.comm_Content,C.comm_Author,C.comm_PostTime,C.blog_ID,L.log_ID,L.log_IsShow FROM blog_Comment AS C,blog_Content AS L WHERE L.log_ID=C.blog_ID ORDER BY comm_PostTime DESC"
		log_LastCommList.Open SQL,Conn,1,1
		SQLQueryNums=SQLQueryNums+1
		If log_LastCommList.EOF AND log_LastCommList.BOF Then
			Redim Arr_LastComm(7,0)
		Else
			Arr_LastComm=log_LastCommList.GetRows
		End If
		log_LastCommList.Close
		Set log_LastCommList=Nothing
		Application.Lock
		Application(CookieName&"_blog_LastComm")=Arr_LastComm
		Application.UnLock
	Else
		Arr_LastComm=Application(CookieName&"_blog_LastComm")
	End IF
	Dim blog_LastCommListNums,blog_LastCommListNumI
	blog_LastCommListNums=Ubound(Arr_LastComm,2)
	For blog_LastCommListNumI=0 To blog_LastCommListNums
		  IF Arr_LastComm(6,blog_LastCommListNumI)=Empty Then
				Response.Write("隐藏日志的评论")
		  Else
				IF DelQuote(Arr_LastComm(1,blog_LastCommListNumI))<>Empty Then
					Dim blog_LastCommContent
					blog_LastCommContent=DelQuote(HTMLEncode(Arr_LastComm(1,blog_LastCommListNumI)))
					Response.Write("<a href=""blogview.asp?logID="&Arr_LastComm(4,blog_LastCommListNumI)&"#commmark_"&Arr_LastComm(0,blog_LastCommListNumI)&""" title="""&Arr_LastComm(2,blog_LastCommListNumI)&" 于 "&DateToStr(Arr_LastComm(3,blog_LastCommListNumI),"Y-m-d H:I A")&" 发表评论：&#13;&#10;"&Replace(Left(blog_LastCommContent,255),"<br>","&#13;&#10;")&""">"&SplitLines(cutStr(Replace(blog_LastCommContent,"<br>",""),20),0)&"</a>")
				Else
					Response.Write("<a href=""blogview.asp?logID="&Arr_LastComm(4,blog_LastCommListNumI)&"#commmark_"&Arr_LastComm(0,blog_LastCommListNumI)&""">没有评论内容，只是引用</a>")
				End IF
		  End IF
		  Response.Write("<br>")
	Next
	Response.Write("</div>")
End Sub

Sub NewBlogList '最新Bolg列表
  Response.Write("<div class=""siderbar_head""><img src=""images/news_types.gif"" border=""0"" align=""absmiddle"" /> 最新日志</div><div class=""siderbar_main"">")
  Dim blog_Topicnewlist
  Set blog_Topicnewlist=Conn.Execute("SELECT TOP 8 T.log_ID,T.log_Title,T.log_Author,log_Content,log_IsShow FROM blog_Content T,blog_Category C WHERE T.log_cateID=C.cate_ID ORDER BY T.log_PostTime DESC")
  IF blog_Topicnewlist.EOF AND blog_Topicnewlist.BOF Then
	  Response.Write("暂无")
  Else
	  Do While NOT blog_Topicnewlist.EOF
	  IF blog_Topicnewlist("log_IsShow")=False Then
		Response.Write("<a href='blogview.asp?logID="&blog_Topicnewlist("log_ID")&"' title='隐藏日志，不显示预览'>"&HTMLEncode(cutStr(blog_Topicnewlist("log_Title"),18))&"</a><br>")
	  Else
      Response.Write("<a href='blogview.asp?logID="&blog_Topicnewlist("log_ID")&"' title='"&HTMLEncode(blog_Topicnewlist("log_Author"))&": "&Replace(HTMLEncode(cutStr(blog_Topicnewlist("log_Content"),200)),"<br>","&#13;&#10;")&"'>"&HTMLEncode(cutStr(blog_Topicnewlist("log_Title"),18))&"</a><br>")
	  End IF
		  blog_Topicnewlist.MoveNext
	  Loop
  End IF
  Set blog_Topicnewlist=Nothing
Response.Write("</div>")
End Sub

Sub hotBlogList '最热Bolg列表
  Response.Write("<div class=""siderbar_head""><img src=""images/sider_hot.gif"" border=""0"" align=""absmiddle"" /> 最热日志</div><div class=""siderbar_main"">")
  Dim blog_hotlist
  Set blog_hotlist=Conn.Execute("SELECT TOP 8 T.log_ID,T.log_Title,T.log_Author,log_Content,log_IsShow,log_ViewNums FROM blog_Content T,blog_Category C WHERE T.log_cateID=C.cate_ID ORDER BY T.log_ViewNums DESC")
  IF blog_hotlist.EOF AND blog_hotlist.BOF Then
	  Response.Write("暂无最热日志")
  Else
	  Do While NOT blog_hotlist.EOF
	  IF blog_hotlist("log_IsShow")=False Then
		Response.Write("<a href='blogview.asp?logID="&blog_hotlist("log_ID")&"' title='隐藏日志，不显示预览'>"&HTMLEncode(cutStr(blog_hotlist("log_Title"),18))&"</a><br>")
	  Else
      Response.Write("<a href='blogview.asp?logID="&blog_hotlist("log_ID")&"' title='"&HTMLEncode(blog_hotlist("log_Author"))&": "&Replace(HTMLEncode(cutStr(blog_hotlist("log_Content"),200)),"<br>","&#13;&#10;")&"'>"&HTMLEncode(cutStr(blog_hotlist("log_Title"),18))&"</a><br>")
	  End IF
		  blog_hotlist.MoveNext
	  Loop
  End IF
  Set blog_hotlist=Nothing
Response.Write("</div>")
End Sub

Sub SiteInfo '站点统计
		  Response.Write("<div class=""siderbar_head""><img src=""images/sider_siteinfo.gif"" border=""0"" align=""absmiddle"" /> 站点统计</div><div class=""siderbar_main"">日志："&blog_LogNums&" 篇<br><a href=""commlist.asp"">评论："&blog_CommNums&" 篇</a><br><a href=""tblist.asp"">引用："&blog_QuoteNums&" 个</a><br><a href=""member.asp"">会员："&blog_MemNums&" 人</a><br><a href=""download.asp"">资源："&blog_downloadNums&" 个</a><br><a href=""photo.asp"">相册："&blog_photoNums&" 张</a><br><a href=""forumview.asp"">主题："&blog_ThreadNums&" 个</a><br><a href=""forumview.asp"">帖子："&blog_PostNums&" 个</a><br><a href=""guestbook.asp"">留言："&blog_GuestbookNums&" 个</a><br><a href=""blogvisit.asp?tjtype=5"">在线："&Application(CookieName & "_blog_displayOnlineTitle2")&"</a><br><a href=""blogvisit.asp"">访问："&blog_VisitNums&" 次</a><br>建立：2006-01-22</div>")
		  
End Sub

Sub CategoryList '分类列表
	Dim blog_CategoryNums,blog_CategoryNumI
	blog_CategoryNums=Ubound(Arr_Category,2)
	For blog_CategoryNumI=0 To blog_CategoryNums
		Response.Write("&nbsp;<img src=""images/icon_gray.gif"">&nbsp;<a href=""default.asp?cateID="&Arr_Category(0,blog_CategoryNumI)&""">"&Arr_Category(1,blog_CategoryNumI)&"</a>")
	Next
End Sub

Sub blogSearch '站点搜索
	Response.Write("<div class=""siderbar_head""><img src=""images/sider_search.gif"" border=""0"" align=""absmiddle""> 日志搜索</div><div class=""siderbar_main""><form name=""blogsearch"" method=""post"" action=""search.asp""><input name=""SearchContent"" type=""text"" id=""SearchContent"" size=""18"" title=""请输入要搜索的内容"" /> <input name=""Submit"" type=""Image"" id=""Submit"" value="""" src=""images/go.gif"" align=""absmiddle"" style=""height:17px;width:18px"" /><br><span class=""smalltxt""><input type=""checkbox"" name=""Is_Title"" value=""1"" style=""height:14px;width:14px"" checked>&nbsp;标题&nbsp;&nbsp;<input type=""checkbox"" name=""Is_Content"" value=""1"" style=""height:14px;width:14px"">&nbsp;内容</span></form></div>")
End Sub

Sub ShowEnglish
Dim ConnEnglish,EngRs,Sql,EngStr,ChStr,linecount
 Set ConnEnglish= Server.CreateObject("ADODB.Connection")
 ConnEnglish.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("English.mdb")
 ConnEnglish.Open
linecount=20
Sql="select top "&linecount&" * from sentence where"
dim RNumber,i
Randomize
for i=0 to linecount-1
  RNumber = Int((3747) * Rnd + 1)
 if i=0 then 
        Sql=Sql & " id=" & RNumber
 else
        Sql=Sql & " or id=" & RNumber
 end if
next
set EngRs=ConnEnglish.ExeCute(Sql)
Response.Write("<div id=SEng class=""Econtent""><img &#115;rc=""images/zd.gif"" align=""absmiddle""> <span class=EnText></span><div class=ChText></div></div>")
Response.Write("<div style=""height:5px;overflow:hidden""></div>")
response.write ("<script>")
i=0
Do Until EngRs.EOF
EngStr=Replace(EngRs("English"),"'","\'")
response.write ("En["&i&"]='"&EngStr&"'"&chr(10))
response.write ("Ch["&i&"]='"&EngRs("Chinese")&"'"&chr(10))
EngRs.MoveNext
i=i+1
loop
response.write ("showE(1);window.setInterval (""showE()"",30000)</script>")
end sub

Sub ForumList(ForumID)
	Response.Write("<div class=""siderbar_head""><img src=""images/sider_FCATE.gif"" border=""0"" align=""absmiddle"" /> 论坛列表</div><div class=""siderbar_main"">")
	Dim blog_ForumListNumS,blog_ForumListNumI
	blog_ForumListNumS=Ubound(Arr_Forums,2)
	For blog_ForumListNumI=0 To blog_ForumListNumS
		Response.Write("<div>&nbsp;<img src=""images/forums/"&Arr_Forums(0,blog_ForumListNumI)&".gif"" align=""absmiddle"" border=""0"">&nbsp;")
		If Int(Arr_Forums(0,blog_ForumListNumI))=Int(ForumID) Then
			Response.Write("<a href=""forumview.asp?forumID="&Arr_Forums(0,blog_ForumListNumI)&"""><b>"&Arr_Forums(1,blog_ForumListNumI)&"</b></a>")
		Else
			Response.Write("<a href=""forumview.asp?forumID="&Arr_Forums(0,blog_ForumListNumI)&""">"&Arr_Forums(1,blog_ForumListNumI)&"</a>")
		End If
		Response.Write("&nbsp;("&Arr_Forums(3,blog_ForumListNumI)&")</div>")
	Next
Response.Write("</div>")
End Sub

Function ForumTitle(ForumID)
	Dim blog_ForumListNumS,blog_ForumListNumI
	blog_ForumListNumS=Ubound(Arr_Forums,2)
	For blog_ForumListNumI=0 To blog_ForumListNumS
		If Int(Arr_Forums(0,blog_ForumListNumI))=Int(ForumID) Then
			ForumTitle=" - <a href=""forumview.asp?forumID="&Arr_Forums(0,blog_ForumListNumI)&""">"&Arr_Forums(1,blog_ForumListNumI)&"</a>"
		End If
	Next
End Function

sub N_Forum '最新帖子
	Dim N_Forum,Num_Forum
        Response.Write("<div class=""siderbar_head""><img src=""images/sider_Nforum.gif"" border=""0"" align=""absmiddle"" /> 最新帖子</div><div class=""siderbar_main"">")
	Set N_Forum=Conn.ExeCute("SELECT Top 8 thread_ForumID,thread_ID,thread_Author,thread_LastPoster,thread_Icon,thread_IsTop,thread_IsDigest,thread_IsClose,thread_Title,thread_PostNums,thread_ViewNums,thread_LastPost,thread_PostTime FROM blog_Threads ORDER BY thread_PostTime DESC")
        IF N_Forum.EOF AND N_Forum.BOF Then
                Response.Write("暂无最新帖子")
        Else
            Num_Forum=0
            Do While NOT N_Forum.EOF
        	If Num_Forum<8 Then
        	Response.Write("<A HREF='threadview.asp?forumID="&N_Forum("thread_forumID")&"&threadID="&N_Forum("thread_ID")&"' title='帖子主题："&N_Forum("thread_Title")&"&#13;发  帖  人："&N_Forum("thread_Author")&"&#13;浏览次数："&N_Forum("thread_ViewNums")&"&#13;最后回复："&N_Forum("thread_LastPoster")&"&#13;发帖时间："&DateToStr(N_Forum("thread_PostTime"),"Y-m-d-s")&"' target='_blank'>"&HTMLEncode(cutStr(N_Forum("thread_Title"),20))&"</a><br>")
                End If
                N_Forum.MoveNext
            Num_Forum=Num_Forum+1
            Loop
        End IF
        Set N_Forum=Nothing
		Response.Write("</div>")
End Sub

sub T_Forum '最热帖子
	Dim T_Forum,Tum_Forum
        Response.Write("<div class=""siderbar_head""><img src=""images/sider_Tforum.gif"" border=""0"" align=""absmiddle"" /> 最热帖子</div><div class=""siderbar_main"">")
	Set T_Forum=Conn.ExeCute("SELECT Top 5 thread_ForumID,thread_ID,thread_Author,thread_LastPoster,thread_Icon,thread_IsTop,thread_IsDigest,thread_IsClose,thread_Title,thread_PostNums,thread_ViewNums,thread_LastPost,thread_PostTime FROM blog_Threads ORDER BY thread_ViewNums DESC")
        IF T_Forum.EOF AND T_Forum.BOF Then
                Response.Write("暂无最热帖子")
        Else
            Tum_Forum=0
            Do While NOT T_Forum.EOF
        	If Tum_Forum<5 Then
        	Response.Write("<A HREF='threadview.asp?forumID="&T_Forum("thread_forumID")&"&threadID="&T_Forum("thread_ID")&"' title='帖子主题："&T_Forum("thread_Title")&"&#13;发  帖  人："&T_Forum("thread_Author")&"&#13;浏览次数："&T_Forum("thread_ViewNums")&"&#13;最后回复："&T_Forum("thread_LastPoster")&"&#13;发帖时间："&DateToStr(T_Forum("thread_PostTime"),"Y-m-d-s")&"' target='_blank'>"&HTMLEncode(cutStr(T_Forum("thread_Title"),20))&"</a><br>")
                End If
                T_Forum.MoveNext
            Tum_Forum=Tum_Forum+1
            Loop
        End IF
        Set T_Forum=Nothing
		Response.Write("</div>")
End Sub

sub R_Forum '最新回复
	Dim R_Forum,Rum_Forum
        Response.Write("<div class=""siderbar_head""><img src=""images/sider_Rforum.gif"" border=""0"" align=""absmiddle"" /> 最新回复</div><div class=""siderbar_main"">")
	Set R_Forum=Conn.ExeCute("SELECT Top 5 post_ID,post_ThreadID,post_ForumID,post_Content,post_IsTop,post_Author,post_PostTime FROM blog_Posts ORDER BY post_PostTime DESC")
        IF R_Forum.EOF AND R_Forum.BOF Then
                Response.Write("暂无最新回复")
        Else
            Rum_Forum=0
            Do While NOT R_Forum.EOF
        	If Rum_Forum<5 Then
        	Response.Write("<A HREF='threadview.asp?forumID="&R_Forum("post_ForumID")&"&threadID="&R_Forum("post_ThreadID")&"' title='回复内容："&HTMLEncode(cutStr(R_Forum("post_Content"),150))&"&#13;回  复  人："&R_Forum("post_Author")&"&#13;回复时间："&DateToStr(R_Forum("post_PostTime"),"Y-m-d-s")&"' target='_blank'>"&HTMLEncode(cutStr(R_Forum("post_Content"),20))&"</a><br>")
                End If
                R_Forum.MoveNext
            Rum_Forum=Rum_Forum+1
            Loop
        End IF
        Set R_Forum=Nothing
		Response.Write("</div>")
End Sub

Sub forumSearch '论坛搜索
	Response.Write("<div class=""siderbar_head""><img src=""images/sider_search.gif"" border=""0"" align=""absmiddle""> 论坛搜索</div><div class=""siderbar_main""><form name=""forumsearch"" method=""post"" action=""forumsearch.asp""><input name=""SearchContent"" type=""text"" id=""SearchContent"" size=""18"" title=""请输入要搜索的内容"" /> <input name=""Submit"" type=""Image"" id=""Submit"" value="""" src=""images/go.gif"" align=""absmiddle"" style=""height:17px;width:18px"" /></form></div>")
End Sub
%>