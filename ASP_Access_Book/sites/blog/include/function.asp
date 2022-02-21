<%Function IsvalidFile(File_Type)   '限制上传文件类型
	IsvalidFile = False
	Dim GName
	For Each GName in UP_FileType
		If File_Type = GName Then
			IsvalidFile = True
			Exit For
		End If
	Next
End Function

Function IsInteger(Para) '检测是否有效的数字
	IsInteger=False
	If Not (IsNull(Para) Or Trim(Para)="" Or Not IsNumeric(Para)) Then
		IsInteger=True
	End If
End Function

Function CheckLinkStr(Str)
	Str = Replace(Str, "document.cookie", ".")
	Str = Replace(Str, "document.write", ".")
	Str = Replace(Str, "javascript:", "javascript ")
	Str = Replace(Str, "vbscript:", "vbscript ")
	Str = Replace(Str, "javascript :", "javascript ")
	Str = Replace(Str, "vbscript :", "vbscript ")
	Str = Replace(Str, "[", "&#91;")
	Str = Replace(Str, "]", "&#93;")
	Str = Replace(Str, "<", "&#60;")
	Str = Replace(Str, ">", "&#62;")
	Str = Replace(Str, "{", "&#123;")
	Str = Replace(Str, "}", "&#125;")
	Str = Replace(Str, "|", "&#124;")
	Str = Replace(Str, "script", "&#115;cript")
	Str = Replace(Str, "SCRIPT", "&#083;CRIPT")
	Str = Replace(Str, "Script", "&#083;cript")
	Str = Replace(Str, "script", "&#083;cript")
	Str = Replace(Str, "object", "&#111;bject")
	Str = Replace(Str, "OBJECT", "&#079;BJECT")
	Str = Replace(Str, "Object", "&#079;bject")
	Str = Replace(Str, "object", "&#079;bject")
	Str = Replace(Str, "applet", "&#097;pplet")
	Str = Replace(Str, "APPLET", "&#065;PPLET")
	Str = Replace(Str, "Applet", "&#065;pplet")
	Str = Replace(Str, "applet", "&#065;pplet")
	Str = Replace(Str, "embed", "&#101;mbed")
	Str = Replace(Str, "EMBED", "&#069;MBED")
	Str = Replace(Str, "Embed", "&#069;mbed")
	Str = Replace(Str, "embed", "&#069;mbed")
	Str = Replace(Str, "document", "&#100;ocument")
	Str = Replace(Str, "DOCUMENT", "&#068;OCUMENT")
	Str = Replace(Str, "Document", "&#068;ocument")
	Str = Replace(Str, "document", "&#068;ocument")
	Str = Replace(Str, "cookie", "&#099;ookie")
	Str = Replace(Str, "COOKIE", "&#067;OOKIE")
	Str = Replace(Str, "Cookie", "&#067;ookie")
	Str = Replace(Str, "cookie", "&#067;ookie")
	Str = Replace(Str, "event", "&#101;vent")
	Str = Replace(Str, "EVENT", "&#069;VENT")
	Str = Replace(Str, "Event", "&#069;vent")
	Str = Replace(Str, "event", "&#069;vent")
	CheckLinkStr = Str
End Function

Function CheckStr(byVal ChkStr) '检查无效字符
	Dim Str:Str=ChkStr
	Str=Trim(Str)
	If IsNull(Str) Then
		CheckStr = ""
		Exit Function 
	End If
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="(\r\n){3,}"
	Str=re.Replace(Str,"$1$1$1")
	Set re=Nothing
	Str = Replace(Str,"'","''")
	Str = Replace(Str, "select", "sel&#101;ct")
	Str = Replace(Str, "join", "jo&#105;n")
	Str = Replace(Str, "union", "un&#105;on")
	Str = Replace(Str, "where", "wh&#101;re")
	Str = Replace(Str, "insert", "ins&#101;rt")
	Str = Replace(Str, "delete", "del&#101;te")
	Str = Replace(Str, "update", "up&#100;ate")
	Str = Replace(Str, "like", "lik&#101;")
	Str = Replace(Str, "drop", "dro&#112;")
	Str = Replace(Str, "create", "cr&#101;ate")
	Str = Replace(Str, "modify", "mod&#105;fy")
	Str = Replace(Str, "rename", "ren&#097;me")
	Str = Replace(Str, "alter", "alt&#101;r")
	Str = Replace(Str, "cast", "ca&#115;t")

	CheckStr=Str
End Function

Function UnCheckStr(Str)
		Str = Replace(Str, "sel&#101;ct", "select")
		Str = Replace(Str, "jo&#105;n", "join")
		Str = Replace(Str, "un&#105;on", "union")
		Str = Replace(Str, "wh&#101;re", "where")
		Str = Replace(Str, "ins&#101;rt", "insert")
		Str = Replace(Str, "del&#101;te", "delete")
		Str = Replace(Str, "up&#100;ate", "update")
		Str = Replace(Str, "lik&#101;", "like")
		Str = Replace(Str, "dro&#112;", "drop")
		Str = Replace(Str, "cr&#101;ate", "create")
		Str = Replace(Str, "mod&#105;fy", "modify")
		Str = Replace(Str, "ren&#097;me", "rename")
		Str = Replace(Str, "alt&#101;r", "alter")
		Str = Replace(Str, "ca&#115;t", "cast")

		UnCheckStr=Str
End Function

Function HTMLEncode(reString) '转换HTML代码
	Dim Str:Str=reString
	If Not IsNull(Str) Then
		Str = UnCheckStr(Str)
		Str = Replace(Str, "&", "&amp;")
		Str = Replace(Str, ">", "&gt;")
		Str = Replace(Str, "<", "&lt;")
		Str = Replace(Str, CHR(32), "&nbsp;")
	    Str = Replace(Str, CHR(9), "&nbsp;&nbsp;&nbsp;&nbsp;")
		Str = Replace(Str, CHR(9), "&#160;&#160;&#160;&#160;")
		Str = Replace(Str, CHR(34),"&quot;")
		Str = Replace(Str, CHR(39),"&#39;")
		Str = Replace(Str, CHR(13), "")
		Str = Replace(Str, CHR(10), "<br>")
		HTMLEncode = Str
	End If
End Function

Function EditDeHTML(byVal Content)
	EditDeHTML=Content
	IF Not IsNull(EditDeHTML) Then
		EditDeHTML=UnCheckStr(EditDeHTML)
		EditDeHTML=Replace(EditDeHTML,"&","&amp;")
		EditDeHTML=Replace(EditDeHTML,"<","&lt;")
		EditDeHTML=Replace(EditDeHTML,">","&gt;")
		EditDeHTML=Replace(EditDeHTML,CHR(34),"&quot;")
		EditDeHTML=Replace(EditDeHTML,CHR(39),"&#39;")
	End IF
End Function

Function DateToStr(DateTime,ShowType)  '日期转换函数
	Dim DateMonth,DateDay,DateHour,DateMinute
	DateMonth=Month(DateTime)
	DateDay=Day(DateTime)
	DateHour=Hour(DateTime)
	DateMinute=Minute(DateTime)
	If Len(DateMonth)<2 Then DateMonth="0"&DateMonth
	If Len(DateDay)<2 Then DateDay="0"&DateDay
	Select Case ShowType
	Case "Y-m-d"  
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay
	Case "Y-m-d H:I A"
		Dim DateAMPM
		If DateHour>12 Then 
			DateHour=DateHour-12
			DateAMPM="PM"
		Else
			DateHour=DateHour
			DateAMPM="AM"
		End If
		If Len(DateHour)<2 Then DateHour="0"&DateHour	
		If Len(DateMinute)<2 Then DateMinute="0"&DateMinute
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute&" "&DateAMPM
	Case "Y-m-d H:I:S"
		Dim DateSecond
		DateSecond=Second(DateTime)
		If Len(DateHour)<2 Then DateHour="0"&DateHour	
		If Len(DateMinute)<2 Then DateMinute="0"&DateMinute
		If Len(DateSecond)<2 Then DateSecond="0"&DateSecond
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute&":"&DateSecond
	Case "YmdHIS"
		DateSecond=Second(DateTime)
		If Len(DateHour)<2 Then DateHour="0"&DateHour	
		If Len(DateMinute)<2 Then DateMinute="0"&DateMinute
		If Len(DateSecond)<2 Then DateSecond="0"&DateSecond
		DateToStr=Year(DateTime)&DateMonth&DateDay&DateHour&DateMinute&DateSecond	
	Case "ym"
		DateToStr=Right(Year(DateTime),2)&DateMonth
	Case "d"
		DateToStr=DateDay
	Case Else
		If Len(DateHour)<2 Then DateHour="0"&DateHour
		If Len(DateMinute)<2 Then DateMinute="0"&DateMinute
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute
	End Select
End Function

Function IsValidUserName(byVal UserName)
	Dim i,c
	IsValidUserName = True
	For i = 1 To Len(UserName)
		c = Lcase(Mid(UserName, i, 1))
		IF InStr("$!<>?#^%@~`&*(){};:+='"" 		", c) > 0 Then
				IsValidUserName = False
				Exit Function
		End IF
	Next
End Function

Function IsValidEmail(Email) '检测是否有效的E-mail地址
	Dim names, name, i, c
	IsValidEmail = True
	Names = Split(email, "@")
	If UBound(names) <> 1 Then
   		IsValidEmail = False
   		Exit Function
	End If
	For Each name IN names
		If Len(name) <= 0 Then
     		IsValidEmail = False
     		Exit Function
   		End If
   		For i = 1 to Len(name)
     		c = Lcase(Mid(name, i, 1))
     		If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
       			IsValidEmail = false
       			Exit Function
     		End If
   		Next
   		If Left(name, 1) = "." or Right(name, 1) = "." Then
      		IsValidEmail = false
      		Exit Function
   		End If
	Next
	If InStr(names(1), ".") <= 0 Then
   		IsValidEmail = False
   		Exit Function
	End If
	i = Len(names(1)) - InStrRev(names(1), ".")
	If i <> 2 And i <> 3 Then
   		IsValidEmail = False
   		Exit Function
	End If
	If InStr(email, "..") > 0 Then
   		IsValidEmail = False
	End If
End Function

Function MultiPage(Numbers,Perpage,Curpage,Url_Add) '分页函数
	CurPage=Int(Curpage)
	Dim URL
	URL=Request.ServerVariables("Script_Name")&Url_Add
	MultiPage=""
	Dim Page,Offset,PageI
	If Int(Numbers)>Int(PerPage) Then
		Page=10
		Offset=2
		Dim Pages,FromPage,ToPage
		If Numbers Mod Cint(Perpage)=0 Then
			Pages=Int(Numbers/Perpage)
		Else
			Pages=Int(Numbers/Perpage)+1
		End If
		FromPage=Curpage-Offset
		ToPage=Curpage+Page-Offset-1
		If Page>Pages Then
			FromPage=1
			ToPage=Pages
		Else
			If FromPage<1 Then
				Topage=Curpage+1-FromPage
				FromPage=1
				If (ToPage-FromPage)<Page And (ToPage-FromPage)<Pages Then ToPage=Page
			ElseIF Topage>Pages Then
				FromPage =Curpage-Pages +ToPage
				ToPage=Pages
				If (ToPage-FromPage)<Page And (ToPage-FromPage)<Pages Then FromPage=Pages-Page+1
			End If
		End If
		MultiPage="<a href="""&Url&"page=1""><img src=""images/icon_ar.gif"" border=""0"" align=""absmiddle""></a> "
		For PageI=FromPage TO ToPage
			If PageI<>CurPage Then
				MultiPage=MultiPage&"<a href="""&Url&"page="&PageI&""">["&PageI&"]</a>&nbsp;"
			Else
				MultiPage=MultiPage&"<b>["&PageI&"]</b>&nbsp;"
			End If
		Next
		If Int(Pages)>Int(Page) Then
			MultiPage=MultiPage&" ... <a href="""&Url&"page="&Pages&"""> ["&pages&"] <img src=""images/icon_al.gif"" border=""0"" align=""absmiddle""></a>&nbsp;<input type=""text"" name=""custompage"" size=""1"" class=""custompage"" onKeyDown=""javascript: if(window.event.keyCode == 13) window.location='"&Url&"page='+this.value;"">"
		Else
			MultiPage=MultiPage&" <a href="""&Url&"page="&Pages&"""><img src=""images/icon_al.gif"" border=""0"" align=""absmiddle""></a>"
		End If
	End If
End Function

Function SplitLines(byVal Content,byVal ContentNums) '切割内容
	Dim ts,i,l
	If IsNull(Content) Then Exit Function
	i=1
	ts = 0
	For i=1 to Len(Content)
      	l=Mid(Content,i,4)
      	If l="<br>" Then
         	ts=ts+1
      	End If
      	If ts>ContentNums Then Exit For 
	Next
	If ts>ContentNums Then
    	Content=Left(Content,i-1)
	End If
	SplitLines=Content
End Function

Sub GetCurPage
	If CheckStr(Request.QueryString("Page"))<>Empty Then
		Curpage=CheckStr(Request.QueryString("Page"))
		If IsInteger(Curpage)=False Then
			CurPage=1
		Else
			If CurPage<0 Then
				CurPage=1
			End If
		End If
	Else
		CurPage=1
	End If
End Sub

Function Generator(Length)
	Dim i, tempS
	tempS = "abcdefghijklmnopqrstuvwxyz1234567890" 
	Generator = ""
	If isNumeric(Length) = False Then 
		Exit Function 
	End If 
	For i = 1 to Length 
		Randomize 
		Generator = Generator & Mid(tempS,Int((Len(tempS) * Rnd) + 1),1)
	Next 
End Function 

Function CutStr(byVal Str,byVal StrLen)
	Dim l,t,c,i
	l=Len(str)
	t=0
	For i=1 To l
		c=AscW(Mid(str,i,1))
		If c<0 Or c>255 Then t=t+2 Else t=t+1
		IF t>=StrLen Then
			CutStr=left(Str,i)&"..."
			Exit For
		Else
			CutStr=Str
		End If
	Next
End Function

Function Trackback(trackback_url, url, title, excerpt, blog_name) 
	Dim query_string, objXMLHTTP, objDOM
	title = cutStr(Server.URLEncode(title),100)
	excerpt = cutStr(Server.URLEncode(excerpt), 252)
	url = Server.URLEncode(url)
	blog_name = Server.URLEncode(blog_name)
	query_string = "title="&title&"&url="&url&"&blog_name="&blog_name&"&excerpt="&excerpt

	Set objXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
	Set objDom = Server.CreateObject("Microsoft.XMLDOM")

	objXMLHTTP.Open "POST", trackback_url, false
	objXMLHTTP.setRequestHeader "Content-Type","application/x-www-Form-urlencoded"

	'HAndling timeout
	On Error Resume Next
	
	objXMLHTTP.SEnd query_string

	If objXMLHTTP.readyState <> 4 Then
		objXMLHTTP.waitForResponse 15
	End If

	If Err.Number <> 0 Then
		Trackback	= "0$$TrackBack 错误：无法连接服务器"
	Else
		If (objXMLHTTP.readyState <> 4) Or (objXMLHTTP.Status <> 200) Then
			objXMLHTTP.Abort
			Trackback	= "0$$Trackback 超时"
		Else
			objDom.async=false
			objDom.loadXML(objXMLHTTP.responseText) 
			If objDom.parseError.errorCode <> 0 Then
				Trackback	= "0$$TrackBack 响应解析错误"
			Else
				If objDom.getElementsByTagName("error")(0).Text="0" Then
					Trackback	= "1$$Trackback 成功"
				Else
					Trackback	= "0$$Trackback 错误："&objDom.getElementsByTagName("message")(0).Text
				End If
			End If
		End If
	End If

	Set objXMLHTTP = Nothing
	Set objDom = Nothing

End Function

Function DelQuote(strContent)
	If IsNull(strContent) Then Exit Function
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="(\[quote\])(.*?)(\[\/quote\])"
	strContent= re.Replace(strContent,"")
	Set re=Nothing
	DelQuote=strContent
End Function

Function CheckWordFilter(byVal Str)
	Dim log_WordFilterListNumS,log_WordFilterListNumI
	log_WordFilterListNumS=Ubound(Arr_WordFilter,2)
	For log_WordFilterListNumI=0 To log_WordFilterListNumS
		Str=Replace(Str,Arr_WordFilter(1,log_WordFilterListNumI),Arr_WordFilter(2,log_WordFilterListNumI))
	Next
	CheckWordFilter=Str
End Function

Function UnCheckWordFilter(byVal Str)
	Dim log_WordFilterListNumS,log_WordFilterListNumI
	log_WordFilterListNumS=Ubound(Arr_WordFilter,2)
	For log_WordFilterListNumI=0 To log_WordFilterListNumS
		Str=Replace(Str,Arr_WordFilter(2,log_WordFilterListNumI),Arr_WordFilter(1,log_WordFilterListNumI))
	Next
	UnCheckWordFilter=Str
End Function

'去除非法链接
Function Strurls(str,notes)
    Strurls=ubound(split(LCase(str),notes))
End Function

Function CheckWordFilter(byVal Str)
	Dim log_WordFilterListNumS,log_WordFilterListNumI
	log_WordFilterListNumS=Ubound(Arr_WordFilter,2)
	For log_WordFilterListNumI=0 To log_WordFilterListNumS
		Str=Replace(Str,Arr_WordFilter(1,log_WordFilterListNumI),Arr_WordFilter(2,log_WordFilterListNumI))
	Next
	CheckWordFilter=Str
End Function

Function ThreadPage(Numbers,Perpage,Url_Add)
	Dim URL,CurPage
	CurPage=1
	URL="threadview.asp"&Url_Add
	ThreadPage=""
	Dim Page,Offset,PageI
	If Int(Numbers)>Int(PerPage) Then
		Page=8
		Offset=2
		Dim Pages,FromPage,ToPage
		If Numbers Mod Cint(Perpage)=0 Then
			Pages=Int(Numbers/Perpage)
		Else
			Pages=Int(Numbers/Perpage)+1
		End If
		FromPage=Curpage-Offset
		ToPage=Curpage+Page-Offset-1
		If Page>Pages Then
			FromPage=1
			ToPage=Pages
		Else
			If FromPage<1 Then
				Topage=Curpage+1-FromPage
				FromPage=1
				If (ToPage-FromPage)<Page And (ToPage-FromPage)<Pages Then ToPage=Page
			ElseIF Topage>Pages Then
				FromPage =Curpage-Pages +ToPage
				ToPage=Pages
				If (ToPage-FromPage)<Page And (ToPage-FromPage)<Pages Then FromPage=Pages-Page+1
			End If
		End If
		For PageI=FromPage TO ToPage
			ThreadPage=ThreadPage&"&nbsp;<a href="""&URL&"page="&PageI&""">"&PageI&"</a>"
		Next
		If Int(Pages)>Int(Page) Then
			ThreadPage=ThreadPage&"&nbsp;...&nbsp;<a href="""&URL&"page="&Pages&""">"&Pages&"</a>"
		End If
	End If
End Function

Function MultiPage_tag(Numbers,Perpage,Curpage,Url_Add) 'TAG列表分页函数
	CurPage=Int(Curpage)
	Dim URL
	URL=Request.ServerVariables("Script_Name")&Url_Add
	MultiPage_tag=""
	Dim Page,Offset,PageI
	If Int(Numbers)>Int(PerPage) Then
		Page=10
		Offset=2
		Dim Pages,FromPage,ToPage
		If Numbers Mod Cint(Perpage)=0 Then
			Pages=Int(Numbers/Perpage)
		Else
			Pages=Int(Numbers/Perpage)+1
		End If
		FromPage=Curpage-Offset
		ToPage=Curpage+Page-Offset-1
		If Page>Pages Then
			FromPage=1
			ToPage=Pages
		Else
			If FromPage<1 Then
				Topage=Curpage+1-FromPage
				FromPage=1
				If (ToPage-FromPage)<Page And (ToPage-FromPage)<Pages Then ToPage=Page
			ElseIF Topage>Pages Then
				FromPage =Curpage-Pages +ToPage
				ToPage=Pages
				If (ToPage-FromPage)<Page And (ToPage-FromPage)<Pages Then FromPage=Pages-Page+1
			End If
		End If
		MultiPage_tag="<a href="""&Url&"sortBy="&sortBy&"&tags="&tag&"&page=1""><img src=""images/icon_ar.gif"" border=""0"" align=""absmiddle""></a> "
		For PageI=FromPage TO ToPage
			If PageI<>CurPage Then
				MultiPage_tag=MultiPage_tag&"<a href="""&Url&"sortBy="&sortBy&"&tags="&tag&"&page="&PageI&""">["&PageI&"]</a>&nbsp;"
			Else
				MultiPage_tag=MultiPage_tag&"<b>["&PageI&"]</b>&nbsp;"
			End If
		Next
		If Int(Pages)>Int(Page) Then
			MultiPage_tag=MultiPage_tag&" ... <a href="""&Url&"page="&Pages&"""> ["&pages&"] <img src=""images/icon_al.gif"" border=""0"" align=""absmiddle""></a>&nbsp;<input type=""text"" name=""custompage"" size=""1"" class=""custompage"" onKeyDown=""javascript: if(window.event.keyCode == 13) window.location='"&Url&"page='+this.value;"">"
		Else
			MultiPage_tag=MultiPage_tag&" <a href="""&Url&"sortBy="&sortBy&"&tags="&tag&"&page="&Pages&"""><img src=""images/icon_al.gif"" border=""0"" align=""absmiddle""></a>"
		End If
	End If
End Function

'显示TAG 2005-7-26
Function ShowTag(blogID,TagMode)
  SQL="SELECT TagsName,Blog_ID From blog_Tag WHERE blog_ID="&blogID&""
  DIM STAG,STARR,STNUM,STI,taglist,tagxg,tagxg_blog,tagxg_num,Noid
  Noid=" log_ID<>" & blogID & " "
  tagxg_blog = ""
  Set STAG=SERVER.CREATEOBJECT("ADODB.RECORDSET")
  STAG.OPEN sql,conn,1,1
  IF STAG.EOF AND STAG.BOF THEN
     Else
     STARR=STAG.GetRows
     STNUM=Ubound(STARR,2)
     For STI=0 To STNUM
       IF TagMode="Edit" then
          IF STI=STNUM Then
             ShowTag = ShowTag & STARR(0,STI)
             Else
             ShowTag = ShowTag & STARR(0,STI) & ";"
          End IF
        ElseIf TagMode="Meta" then
          IF STI=STNUM Then
             ShowTag=ShowTag&STARR(0,STI)
          Else
             ShowTag=ShowTag&STARR(0,STI)&","
          End IF
        Else
        IF ucase(Trim(CheckStr(Trim(Request.QueryString("tags")))))=ucase(Trim(STARR(0,STI))) Then
            taglist="<font color=#ff0000>"&STARR(0,STI)&"</font>"
          Else
            taglist=STARR(0,STI)
        End IF
        
        '显示相关日志与TAG有关 2005-10-30
        '判断非首页
        IF Request.Querystring("logID")<>Empty Then
        Select Case STNUM
               Case 0
                 tagxg_num= 6
               Case 1
                 tagxg_num= 3
               Case 2
                 tagxg_num= 2 
               Case 3
                 tagxg_num= 2
               Case Else
                 tagxg_num= 1
        End Select
        IF STARR(0,STI)=Empty or STARR(0,STI)="" Then
        Else
        SQL="SELECT L.*,C.cate_Name,A.* FROM blog_Content AS L,blog_Category AS C,blog_tag AS A Where C.cate_ID=L.log_CateID AND L.log_ID=A.blog_ID and tagsName = '" & STARR(0,STI) & "' And " & Noid & " ORDER BY log_IsTop ASC,log_ID DESC"
    Set tagxg=Server.CreateObject("adodb.recordset")
        tagxg.Open sql,conn,1,1
        IF tagxg.eof And tagxg.bof Then
           
        Else
           Do while NOT tagxg.eof
              IF STI>STNUM Then
                  tagxg_blog = tagxg_blog & "<a href=""BlogView.asp?logID="&tagxg("log_ID")&""">" & tagxg("log_Title") & "</a>&nbsp;&nbsp;&nbsp;&nbsp;<span class=""date"">"&DateToStr(tagxg("log_PostTime"),"Y-m-d A")&"&nbsp;&nbsp;"&tagxg("cate_Name")&"</span>"
                 Else
                  tagxg_blog = tagxg_blog & "<a href=""BlogView.asp?logID="&tagxg("log_ID")&""">" & tagxg("log_Title") & "</a>&nbsp;&nbsp;&nbsp;&nbsp;<span class=""date"">"&DateToStr(tagxg("log_PostTime"),"Y-m-d A")&"&nbsp;&nbsp;&nbsp;"&tagxg("cate_Name")&"</span><br />"
              End IF
              Noid = Noid & " And log_ID<>" & tagxg("log_ID") & " " 
              tagxg.movenext
           Loop
           
        End IF
        tagxg.Close
        Set tagxg=NoThing
        End IF
        End IF
        '显示相关日志与TAG有关 结束
        
        IF STI=STNUM Then
           ShowTag = ShowTag & "<a href=""BloglistTag.asp?tags="&Server.URLEncode(STARR(0,STI))&""">" & taglist & "</a>"
          Else
           ShowTag = ShowTag & "<a href=""BloglistTag.asp?tags="&Server.URLEncode(STARR(0,STI))&""">" & taglist & "</a>" & " | "
        End IF
        
       End IF
     Next
     IF TagMode="Edit" then
     ElseIF TagMode="Meta" then
        ShowTag = "Tags,"&ShowTag&""
     ElseIF tagxg_blog=Empty or tagxg_blog="" Then
        IF ShowTag<>"" AND ShowTag<>Empty Then
           ShowTag = "Tags：" & ShowTag & ""
        End IF
     Else
        tagxg_blog="<BR>相关日志：<BR>"&tagxg_blog&""
        ShowTag = "Tags：" & ShowTag & "" & tagxg_blog
     End IF
  END IF
  STAG.CLOSE
  SET STAG=NOTHING
End Function


Function Realremark(byVal Str)
	Realremark=Replace(Str,"<a","<a rel=""nofollow""")
End Function

Sub EditTags(log_ID)
                SQL="Select * from blog_tag where blog_id="&log_ID&""
				Set deltag=Server.CreateObject("Adodb.Recordset")
				deltag.OPEN SQL,CONN,1,1
				DO While NOT deltag.Eof
				   conn.execute ("update blog_tags set TagBlogCount=TagBlogCount-1 where TagName='"&deltag("TagsName")&"'")
				   deltag.MoveNext
				LOOP
				deltag.Close
				set deltag=nothing
				conn.execute ("DELETE  from blog_tag where blog_ID="&log_ID&"")
				conn.execute ("DELETE  from blog_tags where TagBlogCount=0")
End Sub

Sub DelTags(blog_ED)
                SQL="Select * from blog_tag where blog_id="&blog_ED&""
				Set deltag=Server.CreateObject("Adodb.Recordset")
				deltag.OPEN SQL,CONN,1,1
				DO While NOT deltag.Eof
				   conn.execute ("update blog_tags set TagBlogCount=TagBlogCount-1 where TagName='"&deltag("TagsName")&"'")
				   deltag.MoveNext
				LOOP
				deltag.Close
				set deltag=nothing
				conn.execute ("DELETE  from blog_tag where blog_ID="&blog_ED&"")
				conn.execute ("DELETE  from blog_tags where TagBlogCount=0")
End Sub

'显示TAGS分类
Sub TagsList(tMode)
         Dim TagRS,MAXTag
		 IF tMode="Hot" Then
		    Sql="Select TOP 20 * from blog_tags order by TagBlogCount desc, CreateDate asc"
		 Else
		    Sql="Select * from blog_tags order by CreateDate desc"
		 End IF
		 Set TagRS = Server.CreateObject("Adodb.RecordSet")
		 TagRS.Open Sql,Conn,1,1
		 IF TagRS.Eof AND TagRS.Bof Then
		    Response.Write ("目前没有 Tags 分类。")
		 Else
		    '得到当前最多日志分类数
			MAXTag=conn.Execute("select top 1 TagBlogCount From Blog_Tags Order By TagBlogCount Desc")(0)
		    Do While Not TagRS.Eof
			   IF TagRS("TagBlogCount")>=cint(MAXTag*0.2) AND TagRS("TagBlogCount")<=cint(MAXTag*0.6) Then
			      Response.Write ("<a href=""BloglistTag.asp?tags="&Server.URLEncode(TagRS("TagName"))&"""><span class=""Tag_size2"" title=""此Tag共有"&TagRS("TagBlogCount")&"篇日志"">" & TagRS("TagName") & "</span></a>")
			   ElseIF TagRS("TagBlogCount")>cint(MAXTag*0.6) AND TagRS("TagBlogCount")<=cint(MAXTag*0.9) Then
			      Response.Write ("<a href=""BloglistTag.asp?tags="&Server.URLEncode(TagRS("TagName"))&"""><span class=""Tag_size3"" title=""此Tag共有"&TagRS("TagBlogCount")&"篇日志"">" & TagRS("TagName") & "</span></a>")
			   ElseIF TagRS("TagBlogCount")>cint(MAXTag*0.9) Then
			      Response.Write ("<a href=""BloglistTag.asp?tags="&Server.URLEncode(TagRS("TagName"))&"""><span class=""Tag_size4"" title=""此Tag共有"&TagRS("TagBlogCount")&"篇日志"">" & TagRS("TagName") & "</span></a>")
			   Else
			      Response.Write ("<a href=""BloglistTag.asp?tags="&Server.URLEncode(TagRS("TagName"))&"""><span class=""Tag_size1"" title=""此Tag共有"&TagRS("TagBlogCount")&"篇日志"">" & TagRS("TagName") & "</span></a>")
			   End IF
			   Response.Write (" | ")
			   TagRS.MoveNext
			Loop
		 End IF
		 TagRS.Close
		 Set TagRS=NoThing
		 End Sub

%>
