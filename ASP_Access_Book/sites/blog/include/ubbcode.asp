<%
Function UBBCodes(strContent,DisSM,DisUBB,DisIMG,AutoURL,AutoKEY,logID)
	If isEmpty(strContent) Or isNull(strContent) Then
        Exit Function
	ElseIF DisUBB=1 or DisUBB=True Then
		strContent=Replace(strContent,"[#splitline#]","")
		UBBCodes=strContent
		Exit Function
	Else
		strContent=Replace(strContent,"[#splitline#]","")
		Dim re, strMatches, strMatch, tmpStr1, tmpStr2, tmpStr3, tmpStr4, RNDStr
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True

		re.Pattern="\[code\](<br>)+"
		strContent=re.Replace(strContent,"[code]")
		re.Pattern="\[quote\](<br>)+"
		strContent=re.Replace(strContent,"[quote]")

		IF AutoURL=1 or AutoURL=True Then
				re.Pattern="([^=\]][\s]*?|^)(https?|ftp|gopher|news|telnet|mms|rtsp)://([a-z0-9/\-_+=.~!%@?#%&;:$\\()|]+)"
				StrContent=re.Replace(StrContent,"$1[url]$2://$3[/url]")
		End IF

		IF Not DisIMG=1 or DisIMG=True Then

			re.Pattern="\[img\](.*?)\[\/img\]"
			Set strMatches=re.Execute(strContent)
			For Each strMatch In strMatches
				tmpStr1=CheckLinkStr(strMatch.SubMatches(0))
				strContent=Replace(strContent,strMatch.Value,"<img src="""&tmpStr1&""" border=""0"" onload=""javascript:DrawImage(this);""  alt=""按此在新窗口打开图片"" onmouseover=""this.style.cursor='hand';"" onclick=""window.open(this.src);"" />")
			Next
			Set strMatches=Nothing

			re.Pattern="\[img=(left|right|center|absmiddle)\](.*?)\[\/img\]"
			Set strMatches=re.Execute(strContent)
			For Each strMatch In strMatches
				tmpStr1=strMatch.SubMatches(0)
				tmpStr2=CheckLinkStr(strMatch.SubMatches(1))
				strContent=Replace(strContent,strMatch.Value,"<img src="""&tmpStr2&""" align="""&tmpStr1&"""  onload=""javascript:DrawImage(this);""  border=""0"" alt=""按此在新窗口打开图片"" onmouseover=""this.style.cursor='hand';"" onclick=""window.open(this.src);"" />")
			Next
			Set strMatches=Nothing

			strContent=replace(strContent,"[swf]","[swf=550,400,NAP]")
			strContent=replace(strContent,"[wmv]","[wmv=550,400,NAP]")
			strContent=replace(strContent,"[wma]","[wma=450,70,NAP]")
			strContent=replace(strContent,"[rm]","[rm=550,400,NAP]")
			strContent=replace(strContent,"[ra]","[ra=450,60,NAP]")

			re.Pattern="\[(swf|wma|wmv|rm|ra|qt)=(\d*?|),(\d*?|),(AP|NAP)\](.*?)\[\/(swf|wma|wmv|rm|ra|qt)\]"
			Set strMatches=re.Execute(strContent)
			For Each strMatch in strMatches
				RNDStr=Int(7999 * Rnd + 2000)
				tmpStr1=CheckLinkStr(strMatch.SubMatches(3))
				tmpStr2="UBBShowObj('"&strMatch.SubMatches(0)&"','OBJ_"&RNDStr&"','"&strMatch.SubMatches(4)&"','"&strMatch.SubMatches(1)&"','"&strMatch.SubMatches(2)&"');"
				tmpStr3=""
				If strMatch.SubMatches(3)="AP" Then tmpStr3="<script type=""text/javascript"">window.attachEvent(""onload"",function (){"&tmpStr2&"})</script>"
				strContent= Replace(strContent,strMatch.Value,tmpStr3&"<div class=""code_head""><input id=""VOBJ_"&RNDStr&""" type=""hidden"" value=""-1"" /><a href=""javascript:"&tmpStr2&"""><img src=""images/icon_media.gif"" alt=""显示影音文件"" align=""absmiddle"" border=""0"" /> 点击显示/隐藏影音文件</a></div><div id=""OBJ_"&RNDStr&""" class=""code_main"">影音源文件地址：<a href="""&strMatch.SubMatches(4)&""" target=""_blank"">"&strMatch.SubMatches(4)&"</a></div>")
			Next
			Set strMatches=Nothing
		End IF

		re.Pattern = "\[url=(.[^\]]*)\](.*?)\[\/url]"
		Set strMatches=re.Execute(strContent)
		For Each strMatch In strMatches
			tmpStr1=CheckLinkStr(strMatch.SubMatches(0))
			tmpStr2=strMatch.SubMatches(1)
			strContent=Replace(strContent,strMatch.Value,"<a target=""_blank"" href="""&tmpStr1&""">"&tmpStr2&"</a>")
		Next
		Set strMatches=Nothing
		re.Pattern = "\[url](.*?)\[\/url]"
		Set strMatches=re.Execute(strContent)
		For Each strMatch In strMatches
			tmpStr1=CheckLinkStr(strMatch.SubMatches(0))
			tmpStr2=CutURL(tmpStr1)
			strContent=Replace(strContent,strMatch.Value,"<a target=""_blank"" href="""&tmpStr1&""">"&tmpStr2&"</a>")
		Next
		Set strMatches=Nothing
		
		
		re.Pattern="\[fly\](.*)\[\/fly\]"
		strContent=re.Replace(strContent,"<marquee width=90% behavior=alternate scrollamount=3>$1</marquee>")
		re.Pattern="\[move\](.*)\[\/move\]"
		strContent=re.Replace(strContent,"<MARQUEE scrollamount=3>$1</marquee>")	
		re.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
		strContent=re.Replace(strContent,"<table width=$1 ><tr><td style=""filter:shadow(color=$2, 			strength=$3)"">$4</td></tr></table>")
		re.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
		strContent=re.Replace(strContent,"<table width=$1 ><tr><td style=""filter:glow(color=$2, 			strength=$3)"">$4</td></tr></table>")
		re.Pattern = "\[email=(.[^\]]*)\](.*?)\[\/email]"
		strContent = re.Replace(strContent,"<a href=""mailto:$1"">$2</a>")
		re.Pattern = "\[email](.*?)\[\/email]"
		strContent = re.Replace(strContent,"<a href=""mailto:$1"">$1</a>")

                strContent = Replace(strContent,"[sub]","<sub>")
		strContent = Replace(strContent,"[/sub]","</sub>")

		strContent = Replace(strContent,"[sup]","<sup>")
		strContent = Replace(strContent,"[/sup]","</sup>")
		strContent = Replace(strContent,"[list]","<ul>")
		strContent = Replace(strContent,"[list=1]","<ol type=""1"">")
		strContent = Replace(strContent,"[list=a]","<ol type=""a"">")
		strContent = Replace(strContent,"[list=A]","<ol type=""A"">")
		strContent = Replace(strContent,"[*]","<li>")
		strContent = Replace(strContent,"[/list]","</ul></ol>")
		
		re.Pattern="\[font=([^<>\]]*?)\](.*?)\[\/font]"
		strContent=re.Replace(strContent,"<font face=""$1"">$2</font>")
		re.Pattern="\[color=([^<>\]]*?)\](.*?)\[\/color]"
		strContent=re.Replace(strContent,"<font color=""$1"">$2</font>")
		re.Pattern="\[align=([^<>\]]*?)\](.*?)\[\/align]"
		strContent=re.Replace(strContent,"<div align=""$1"">$2</div>")
		re.Pattern="\[size=(\d*?)\](.*?)\[\/size]"
		strContent=re.Replace(strContent,"<font size=""$1"">$2</font>")
		re.Pattern="\[b\](.*?)\[\/b]"
		strContent=re.Replace(strContent,"<strong>$1</strong>")	
		re.Pattern="\[i\](.*?)\[\/i]"
		strContent=re.Replace(strContent,"<em>$1</em>")	
		re.Pattern="\[u\](.*?)\[\/u]"
		strContent=re.Replace(strContent,"<u>$1</u>")

                re.Pattern="\[light\](.*?)\[\/light]"
                strContent=re.Replace(strContent,"<span style=""behavior:url(include/font.htc)"">$1</span>")
		
			re.Pattern="\[code\](.*?)\[\/code\]"
		Set strMatches=re.Execute(strContent)
		For Each strMatch In strMatches
			RNDStr=Int(7999 * Rnd + 2000)
			tmpStr1=strMatch.SubMatches(0)
			strContent= Replace(strContent,strMatch.Value,"<script type=""text/javascript"">window.attachEvent(""onload"",function (){AutoSizeDIV('CODE_"&RNDStr&"')})</script><table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""code_head""><tr><td>程序代码：</td><td align=""right""><a href=""javascript:CopyText(document.all.CODE_"&RNDStr&");"">[ 复制代码到剪贴板 ]</a> </td></tr></table><div class=""code_main"" id=""CODE_"&RNDStr&""" style=""overflow-y:auto;overflow-x:hidden;height:150px;"">"&tmpStr1&"</div>")
		Next
		Set strMatches=Nothing


		re.Pattern="\[quote\](.*?)\[\/quote\]"
		strContent= re.Replace(strContent,"<div class=""code_head"">引用内容：</div><div class=""code_main"">$1</div>")

		re.Pattern = "\[down=(.[^\]]*)\](.*?)\[\/down]"
		strContent = re.Replace(strContent,"<img src=""images/download.gif"" align=""absmiddle"" /> <a href=""$1"" target=""_blank"">$2</a>")
        
       re.Pattern = "\[download=(.[^\]]*)\](.*?)\[\/download]" '会员下载
		If memStatus="SupAdmin" OR memStatus="Admin" OR memStatus="Member" Then
		strContent = re.Replace(strContent,"<img src=""images/down.gif"" align=""absmiddle"" /> <a href=""$1"" target=""_blank"">$2</a>")
		Else
		strContent = re.Replace(strContent,"<div class=""down_heads"">以下内容需<font color=#0066CC><a href=""register.asp"">注册成员</font></a>才可以下载!<br /><img src=""images/down.gif"" align=""absmiddle"" /> $2</div>")
		End if

        
        strcontent=pages(strcontent,logID)
        re.pattern = "\[page\](.*?)\[\/page]"		
        strcontent = re.replace(strcontent,"")
		
		IF Not DisSM=1 or DisSM=False Then
			Dim log_SmiliesListNums,log_SmiliesListNumI
			log_SmiliesListNums=Ubound(Arr_Smilies,2)
			For log_SmiliesListNumI=0 To log_SmiliesListNums
				strContent=Replace(strContent,Arr_Smilies(2,log_SmiliesListNumI)," <img src=""images/smilies/"&Arr_Smilies(1,log_SmiliesListNumI)&""" border=""0"" align=""absmiddle"" />")
			Next
		End IF

		IF AutoKEY=1 or AutoKEY=True Then
			Dim log_KeywordsListNums,log_KeywordsListNumI
			log_KeywordsListNums=Ubound(Arr_Keywords,2)
			For log_KeywordsListNumI=0 To log_KeywordsListNums
				IF Arr_Keywords(3,log_KeywordsListNumI)<>Empty Then
					strContent=Replace(strContent,Arr_Keywords(1,log_KeywordsListNumI),"<a href="""&Arr_Keywords(2,log_KeywordsListNumI)&""" target=""_blank""><img src=""images/keywords/"&Arr_Keywords(3,log_KeywordsListNumI)&""" border=""0"" align=""absmiddle""> "&Arr_Keywords(1,log_KeywordsListNumI)&"</a>")
				Else
					strContent=Replace(strContent,Arr_Keywords(1,log_KeywordsListNumI),"<a href="""&Arr_Keywords(2,log_KeywordsListNumI)&""" target=""_blank"">"&Arr_Keywords(1,log_KeywordsListNumI)&"</a>")
				End IF
			Next
		End IF

If memName<>Empty Then
re.Pattern="\[mem\](.*?)\[\/mem\]"
strContent= re.Replace(strContent,"<hr size=1><font color=#9D9D9D>以下是注册会员才能看到的内容</font><BR><BR>$1<BR><BR><hr size=1>")
ELSE
re.Pattern="\[mem\](.*?)\[\/mem\]"
strContent= re.Replace(strContent,"<BR><div class=""code_main""> <b>此部分内容只有注册会员才能看到</b></div><BR>")
END IF


		Set re=Nothing
		   
		   UBBCodes="<p id=fp>"&strContent&"</p>"

	End IF

End Function

function pages(Str,logID)
        Dim i, start, nextstart, t, stat, statt, pp, pagecount, pagesize, articleid 
        Dim pa, articletxt, articletext, contenttext, pagelist 
        contenttext = str

        stat = 0
        statt = 0
        start = 0
        nextstart = 0
        pagesize = 2
        pagecount = 0

        pa = Request("pages")
        If (pa = "" Or pa<>null) Then
            pa = "1"
        End If
        pp = cint(pa)

        articletxt = contenttext

        If (len(articletxt) >= pagesize) Then
            t = len(articletxt) / pagesize

            For i = 0 To t
                If (start + pagesize < len(articletxt)) Then
                    'stat = instr(start + pagesize,articletxt,"chr(13)")

                    'If (stat =null) Then
                    stat = instr(start + pagesize,articletxt,"[/page]")
                    'End If
					'response.write stat
                End If
                If (stat=0) Then
                    
					if i=0 then
					   articletext=articletxt
					end if
					i = t+1
                Else
                    nextstart = stat
                    If (start + pagesize >= len(articletxt)) Then
                        nextstart = len(articletxt)
					End If
                    If (pp = i + 1) Then
					if start=0 then 
					   start=1
					end if
					
                        articletext = mid(articletxt,start,nextstart - start)
                        'Label1.Text = articletext
                    End If
                    start = stat
                    pagecount = pagecount + 1
                End If

            Next
        End If

       If (pagecount > 1) Then
            If (pp - 1 > 0) Then
                pagelist = pagelist & "<a href=?logID=" & logID & "&pages=" & (pp - 1) & "><img src=""images/icon_ar.gif"" border=""0"" alt=""上一页""></a> "
            Else
               IF PP=1 THEN
                pagelist = pagelist & "<img src=""images/icon_ar.gif"" border=""0""> "
               ELSE
                pagelist = pagelist & "<a href=?logID=" & logID & "&pages=" & (1) & "><img src=""images/icon_ar.gif"" border=""0"" alt=""上一页""></a> "
               END IF
            End If
            For i = 1 To pagecount
                If (i = pp) Then
                    pagelist = pagelist & "<span class=""smalltxt""><b>[" & i & "]</b></span> "
                Else
                    pagelist = pagelist & "<a href=?logID=" & logID & "&pages=" & i & "><span class=""smalltxt"">[" & i & "]</span></a> "
                End If
            Next
            If (pp + 1 > pagecount) Then
               IF pp=pagecount THEN
                  pagelist = pagelist & "<img src=""images/icon_al.gif"" border=""0"">"
                ELSE
                  pagelist = pagelist & "<a href=?logID=" & logID & "&pages=" & (pagecount) & "><img src=""images/icon_al.gif"" border=""0"" alt=""下一页""></a></p>"
               END IF
            Else           
                pagelist = pagelist & "<a href=?logID=" & logID & "&pages=" & (pp + 1) & "><img src=""images/icon_al.gif"" border=""0"" alt=""下一页""></a></p>"  
            End If
        End If


		articletext=replace(replace(articletext,"[page]",""),"[/page]","")
		pages=articletext &"<p>"& pagelist
end function

Function UBBCode(strContent,DisSM,DisUBB,DisIMG,AutoURL,AutoKEY)
	If isEmpty(strContent) Or isNull(strContent) Then
        Exit Function
	ElseIF DisUBB=1 or DisUBB=True Then
		strContent=Replace(strContent,"[#splitline#]","")
		UBBCode=strContent
		Exit Function
	Else
		strContent=Replace(strContent,"[#splitline#]","")
		Dim re, strMatches, strMatch, tmpStr1, tmpStr2, tmpStr3, tmpStr4, RNDStr
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True

		re.Pattern="\[code\](<br>)+"
		strContent=re.Replace(strContent,"[code]")
		re.Pattern="\[quote\](<br>)+"
		strContent=re.Replace(strContent,"[quote]")

		IF AutoURL=1 Then
				re.Pattern="([^=\]][\s]*?|^)(https?|ftp|gopher|news|telnet|mms|rtsp)://([a-z0-9/\-_+=.~!%@?#%&;:$\\()|]+)"
				StrContent=re.Replace(StrContent,"$1[url]$2://$3[/url]")
		End IF

		IF Not DisIMG=1 or DisIMG=True Then

			re.Pattern="\[img\](.*?)\[\/img\]"
			Set strMatches=re.Execute(strContent)
			For Each strMatch In strMatches
				tmpStr1=CheckLinkStr(strMatch.SubMatches(0))
				strContent=Replace(strContent,strMatch.Value,"<img src="""&tmpStr1&""" border=""0"" onload=""javascript:DrawImage(this);""  alt=""按此在新窗口打开图片"" onmouseover=""this.style.cursor='hand';"" onclick=""window.open(this.src);"" />")
			Next
			Set strMatches=Nothing

			re.Pattern="\[img=(left|right|center|absmiddle)\](.*?)\[\/img\]"
			Set strMatches=re.Execute(strContent)
			For Each strMatch In strMatches
				tmpStr1=strMatch.SubMatches(0)
				tmpStr2=CheckLinkStr(strMatch.SubMatches(1))
				strContent=Replace(strContent,strMatch.Value,"<img src="""&tmpStr2&""" align="""&tmpStr1&"""  onload=""javascript:DrawImage(this);""  border=""0"" alt=""按此在新窗口打开图片"" onmouseover=""this.style.cursor='hand';"" onclick=""window.open(this.src);"" />")
			Next
			Set strMatches=Nothing

			strContent=replace(strContent,"[swf]","[swf=550,400,NAP]")
			strContent=replace(strContent,"[wmv]","[wmv=550,400,NAP]")
			strContent=replace(strContent,"[wma]","[wma=450,70,NAP]")
			strContent=replace(strContent,"[rm]","[rm=550,400,NAP]")
			strContent=replace(strContent,"[ra]","[ra=450,60,NAP]")

			re.Pattern="\[(swf|wma|wmv|rm|ra|qt)=(\d*?|),(\d*?|),(AP|NAP)\](.*?)\[\/(swf|wma|wmv|rm|ra|qt)\]"
			Set strMatches=re.Execute(strContent)
			For Each strMatch in strMatches
				RNDStr=Int(7999 * Rnd + 2000)
				tmpStr1=CheckLinkStr(strMatch.SubMatches(3))
				tmpStr2="UBBShowObj('"&strMatch.SubMatches(0)&"','OBJ_"&RNDStr&"','"&strMatch.SubMatches(4)&"','"&strMatch.SubMatches(1)&"','"&strMatch.SubMatches(2)&"');"
				tmpStr3=""
				strContent= Replace(strContent,strMatch.Value,tmpStr3&"<div class=""code_head""><input id=""VOBJ_"&RNDStr&""" type=""hidden"" value=""-1"" /><a href=""javascript:"&tmpStr2&"""><img src=""images/icon_media.gif"" alt=""显示影音文件"" align=""absmiddle"" border=""0"" /> 点击显示/隐藏影音文件</a></div><div id=""OBJ_"&RNDStr&""" class=""code_main"">影音源文件地址：<a href="""&strMatch.SubMatches(4)&""" target=""_blank"">"&strMatch.SubMatches(4)&"</a></div>")
			Next
			Set strMatches=Nothing
		End IF

		re.Pattern = "\[url=(.[^\]]*)\](.*?)\[\/url]"
		Set strMatches=re.Execute(strContent)
		For Each strMatch In strMatches
			tmpStr1=CheckLinkStr(strMatch.SubMatches(0))
			tmpStr2=strMatch.SubMatches(1)
			strContent=Replace(strContent,strMatch.Value,"<a target=""_blank"" href="""&tmpStr1&""">"&tmpStr2&"</a>")
		Next
		Set strMatches=Nothing
		re.Pattern = "\[url](.*?)\[\/url]"
		Set strMatches=re.Execute(strContent)
		For Each strMatch In strMatches
			tmpStr1=CheckLinkStr(strMatch.SubMatches(0))
			tmpStr2=CutURL(tmpStr1)
			strContent=Replace(strContent,strMatch.Value,"<a target=""_blank"" href="""&tmpStr1&""">"&tmpStr2&"</a>")
		Next
		Set strMatches=Nothing
		re.Pattern = "\[email=(.[^\]]*)\](.*?)\[\/email]"
		strContent = re.Replace(strContent,"<a href=""mailto:$1"">$2</a>")
		re.Pattern = "\[email](.*?)\[\/email]"
		strContent = re.Replace(strContent,"<a href=""mailto:$1"">$1</a>")

                re.Pattern="\[fly\](.*)\[\/fly\]"
		strContent=re.Replace(strContent,"<marquee width=90% behavior=alternate scrollamount=3>$1</marquee>")
		re.Pattern="\[move\](.*)\[\/move\]"
		strContent=re.Replace(strContent,"<MARQUEE scrollamount=3>$1</marquee>")	
		re.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
		strContent=re.Replace(strContent,"<table width=$1 ><tr><td style=""filter:shadow(color=$2, 			strength=$3)"">$4</td></tr></table>")
		re.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
		strContent=re.Replace(strContent,"<table width=$1 ><tr><td style=""filter:glow(color=$2, 			strength=$3)"">$4</td></tr></table>")

                strContent = Replace(strContent,"[sub]","<sub>")
		strContent = Replace(strContent,"[/sub]","</sub>")

		strContent = Replace(strContent,"[sup]","<sup>")
		strContent = Replace(strContent,"[/sup]","</sup>")
		strContent = Replace(strContent,"[list]","<ul>")
		strContent = Replace(strContent,"[list=1]","<ol type=""1"">")
		strContent = Replace(strContent,"[list=a]","<ol type=""a"">")
		strContent = Replace(strContent,"[list=A]","<ol type=""A"">")
		strContent = Replace(strContent,"[*]","<li>")
		strContent = Replace(strContent,"[/list]","</ul></ol>")
		
		re.Pattern="\[light\](.*?)\[\/light]"
                strContent=re.Replace(strContent,"<span style=""behavior:url(include/font.htc)"">$1</span>")
		
		re.Pattern="\[font=([^<>\]]*?)\](.*?)\[\/font]"
		strContent=re.Replace(strContent,"<font face=""$1"">$2</font>")
		re.Pattern="\[color=([^<>\]]*?)\](.*?)\[\/color]"
		strContent=re.Replace(strContent,"<font color=""$1"">$2</font>")
		re.Pattern="\[align=([^<>\]]*?)\](.*?)\[\/align]"
		strContent=re.Replace(strContent,"<div align=""$1"">$2</div>")
		re.Pattern="\[size=(\d*?)\](.*?)\[\/size]"
		strContent=re.Replace(strContent,"<font size=""$1"">$2</font>")
		re.Pattern="\[b\](.*?)\[\/b]"
		strContent=re.Replace(strContent,"<strong>$1</strong>")	
		re.Pattern="\[i\](.*?)\[\/i]"
		strContent=re.Replace(strContent,"<em>$1</em>")	
		re.Pattern="\[u\](.*?)\[\/u]"
		strContent=re.Replace(strContent,"<u>$1</u>")

			re.Pattern="\[code\](.*?)\[\/code\]"
		Set strMatches=re.Execute(strContent)
		For Each strMatch In strMatches
			RNDStr=Int(7999 * Rnd + 2000)
			tmpStr1=strMatch.SubMatches(0)
			strContent= Replace(strContent,strMatch.Value,"<script type=""text/javascript"">window.attachEvent(""onload"",function (){AutoSizeDIV('CODE_"&RNDStr&"')})</script><table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""code_head""><tr><td>程序代码：</td><td align=""right""><a href=""javascript:CopyText(document.all.CODE_"&RNDStr&");"">[ 复制代码到剪贴板 ]</a> </td></tr></table><div class=""code_main"" id=""CODE_"&RNDStr&""" style=""overflow-y:auto;overflow-x:hidden;height:150px;"">"&tmpStr1&"</div>")
		Next
		Set strMatches=Nothing


		re.Pattern="\[quote\](.*?)\[\/quote\]"
		strContent= re.Replace(strContent,"<div class=""code_head"">引用内容：</div><div class=""code_main"">$1</div>")

		re.Pattern = "\[down=(.[^\]]*)\](.*?)\[\/down]"
		strContent = re.Replace(strContent,"<img src=""images/download.gif"" align=""absmiddle"" /> <a href=""$1"" target=""_blank"">$2</a>")

		 re.Pattern = "\[download=(.[^\]]*)\](.*?)\[\/download]" '会员下载
		If memStatus="SupAdmin" OR memStatus="Admin" OR memStatus="Member" Then
		strContent = re.Replace(strContent,"<img src=""images/down.gif"" align=""absmiddle"" /> <a href=""$1"" target=""_blank"">$2</a>")
		Else
		strContent = re.Replace(strContent,"<div class=""down_heads"">以下内容需<a href=""register.asp""><font color=#0066CC>注册成员</font></a>才可以下载!<br /><img src=""images/down.gif"" align=""absmiddle"" /> $2</div>")
		End if
		IF Not DisSM=1 or DisSM=True Then
			Dim log_SmiliesListNums,log_SmiliesListNumI
			log_SmiliesListNums=Ubound(Arr_Smilies,2)
			For log_SmiliesListNumI=0 To log_SmiliesListNums
				strContent=Replace(strContent,Arr_Smilies(2,log_SmiliesListNumI)," <img src=""images/smilies/"&Arr_Smilies(1,log_SmiliesListNumI)&""" border=""0"" align=""absmiddle"" />")
			Next
		End IF

		IF AutoKEY=1 or AutoURL=True Then
			Dim log_KeywordsListNums,log_KeywordsListNumI
			log_KeywordsListNums=Ubound(Arr_Keywords,2)
			For log_KeywordsListNumI=0 To log_KeywordsListNums
				IF Arr_Keywords(3,log_KeywordsListNumI)<>Empty Then
					strContent=Replace(strContent,Arr_Keywords(1,log_KeywordsListNumI),"<a href="""&Arr_Keywords(2,log_KeywordsListNumI)&""" target=""_blank""><img src=""images/keywords/"&Arr_Keywords(3,log_KeywordsListNumI)&""" border=""0"" align=""absmiddle""> "&Arr_Keywords(1,log_KeywordsListNumI)&"</a>")
				Else
					strContent=Replace(strContent,Arr_Keywords(1,log_KeywordsListNumI),"<a href="""&Arr_Keywords(2,log_KeywordsListNumI)&""" target=""_blank"">"&Arr_Keywords(1,log_KeywordsListNumI)&"</a>")
				End IF
			Next
		End IF
If memName<>Empty Then
re.Pattern="\[mem\](.*?)\[\/mem\]"
strContent= re.Replace(strContent,"<hr size=1><font color=#9D9D9D>以下是注册会员才能看到的内容</font><BR><BR>$1<BR><BR><hr size=1>")
ELSE
re.Pattern="\[mem\](.*?)\[\/mem\]"
strContent= re.Replace(strContent,"<BR><div class=""code_main""> <b>此部分内容只有注册会员才能看到</b></div><BR>")
END IF
		Set re=Nothing
strContent=replace(replace(strContent,"[page]",""),"[/page]","")
		UBBCode="<p id=fp>"&strContent&"</p>"

	End IF

End Function

Function CutURL(URLStr)
	Dim URLLen
	URLLen=Len(URLStr)
	If URLLen>65 Then
		CutURL=Left(URLStr,URLLen*0.5)&" ... "&Right(URLStr,URLLen*0.3)
	Else
		CutURL=URLStr
	End If
End Function
%>