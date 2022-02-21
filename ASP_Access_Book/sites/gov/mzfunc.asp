<%
 function ishandset(handno)
	 ishandset=true
	 if len(handno)<>11 then
	 ishandset=false
	   end if
	   if not isnumeric(handno) then
	   ishandset=false
	   end if
	   if cstr(left(handno,3))<>"13" then
	   ishandset=false
	   end if
 end function
function doCode(fString, fOTag, fCTag, fROTag, fRCTag)
	dim fOTagPos, fCTagPos
	fOTagPos = Instr(1, fString, fOTag, 1)
	fCTagPos = Instr(1, fString, fCTag, 1)
	while (fCTagPos > 0 and fOTagPos > 0)
		fString = replace(fString, fOTag, fROTag, 1, 1, 1)
		fString = replace(fString, fCTag, fRCTag, 1,1, 1)
		fOTagPos = Instr(1, fString, fOTag, 1)
		fCTagPos = Instr(1, fString, fCTag, 1)
	wend
	doCode = fString
end function

function HTMLEncode(fString)
	fString = replace(fString, ">", "&gt;")
	fString = replace(fString, "<", "&lt;")
    fString= replace(fString," ","&nbsp;")
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	HTMLEncode = fString
end function

function UBBCode(strContent)
	'on error resume next
	strContent = HTMLEncode(strContent)
	dim objRegExp
	Set objRegExp=new RegExp
	objRegExp.IgnoreCase =true
	objRegExp.Global=True

	objRegExp.Pattern="(\[URL\])(.[^\[]*)(\[\/URL\])"
	strContent= objRegExp.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$2</A>")

	objRegExp.Pattern="(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
	strContent= objRegExp.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$3</A>")

	objRegExp.Pattern="(\[EMAIL\])(.[^\[]*)(\[\/EMAIL\])"
	strContent= objRegExp.Replace(strContent,"<A HREF=""mailto:$2"">$2</A>")
	objRegExp.Pattern="(\[EMAIL=(.[^\[]*)\])(.[^\[]*)(\[\/EMAIL\])"
	strContent= objRegExp.Replace(strContent,"<A HREF=""mailto:$2"" TARGET=_blank>$3</A>")

	objRegExp.Pattern="(\[FLASH\])(.[^\[]*)(\[\/FLASH\])"
	strContent= objRegExp.Replace(strContent,"<OBJECT  width=500 height=400><PARAM NAME=movie VALUE=""$2""><PARAM NAME=quality VALUE=high><embed src=""$2"" quality=high   width=500 height=400>$2</embed></OBJECT>")

	objRegExp.Pattern="(\[IMG\])(.[^\[]*)(\[\/IMG\])"
	strContent=objRegExp.Replace(strContent,"<IMG SRC=""$2"" border=0>")

	objRegExp.Pattern="(\[HTML\])(.[^\[]*)(\[\/HTML\])"
	strContent=objRegExp.Replace(strContent,"<SPAN><p><IMG src=pic/code.gif align=absBottom>该篇文章附带的 HTML 代码片段如下:<BR><TEXTAREA style=""WIDTH: 94%; BACKGROUND-COLOR: #f7f7f7"" name=textfield rows=10>$2</TEXTAREA><BR><INPUT onclick=runEx() type=button value=运行此代码 name=Button> [Ctrl+A 全部选择   提示:你可先修改部分代码，再按运行]</SPAN><BR>")

	objRegExp.Pattern="(\[color=(.[^\[]*)\])(.[^\[]*)(\[\/color\])"
	strContent=objRegExp.Replace(strContent,"<font color=$2>$3</font>")
	objRegExp.Pattern="(\[face=(.[^\[]*)\])(.[^\[]*)(\[\/face\])"
	strContent=objRegExp.Replace(strContent,"<font face=$2>$3</font>")
	objRegExp.Pattern="(\[align=(.[^\[]*)\])(.[^\[]*)(\[\/align\])"
	strContent=objRegExp.Replace(strContent,"<div align=$2>$3</div>")

	objRegExp.Pattern="(\[QUOTE\])(.[^\[]*)(\[\/QUOTE\])"
	strContent=objRegExp.Replace(strContent,"<BLOCKQUOTE><font size=1 face=""Verdana, Arial"">quote:</font><HR>$2<HR></BLOCKQUOTE>")
	objRegExp.Pattern="(\[fly\])(.[^\[]*)(\[\/fly\])"
	strContent=objRegExp.Replace(strContent,"<marquee width=90% behavior=alternate scrollamount=3>$2</marquee>")
	objRegExp.Pattern="(\[move\])(.[^\[]*)(\[\/move\])"
	strContent=objRegExp.Replace(strContent,"<MARQUEE scrollamount=3>$2</marquee>")
	objRegExp.Pattern="(\[glow=(.[^\[]*),(.[^\[]*),(.[^\[]*)\])(.[^\[]*)(\[\/glow\])"
	strContent=objRegExp.Replace(strContent,"<table width=$2 style=""filter:glow(color=$3, strength=$4)"">$5</table>")
	objRegExp.Pattern="(\[SHADOW=(.[^\[]*),(.[^\[]*),(.[^\[]*)\])(.[^\[]*)(\[\/SHADOW\])"
	strContent=objRegExp.Replace(strContent,"<table width=$2 style=""filter:shadow(color=$3, direction=$4)"">$5</table>")
    
	objRegExp.Pattern="(\[i\])(.[^\[]*)(\[\/i\])"
	strContent=objRegExp.Replace(strContent,"<i>$2</i>")
	objRegExp.Pattern="(\[u\])(.[^\[]*)(\[\/u\])"
	strContent=objRegExp.Replace(strContent,"<u>$2</u>")
	objRegExp.Pattern="(\[b\])(.[^\[]*)(\[\/b\])"
	strContent=objRegExp.Replace(strContent,"<b>$2</b>")
	objRegExp.Pattern="(\[fly\])(.[^\[]*)(\[\/fly\])"
	strContent=objRegExp.Replace(strContent,"<marquee>$2</marquee>")

	objRegExp.Pattern="(\[size=1\])(.[^\[]*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=1>$2</font>")
	objRegExp.Pattern="(\[size=2\])(.[^\[]*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=2>$2</font>")
	objRegExp.Pattern="(\[size=3\])(.[^\[]*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=3>$2</font>")
	objRegExp.Pattern="(\[size=4\])(.[^\[]*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=4>$2</font>")

	strContent = doCode(strContent, "[list]", "[/list]", "<ul>", "</ul>")
	strContent = doCode(strContent, "[list=1]", "[/list]", "<ol type=1>", "</ol id=1>")
	strContent = doCode(strContent, "[list=a]", "[/list]", "<ol type=a>", "</ol id=a>")
	strContent = doCode(strContent, "[*]", "[/*]", "<li>", "</li>")
	strContent = doCode(strContent, "[code]", "[/code]", "<pre id=code><font size=1 face=""Verdana, Arial"" id=code>", "</font id=code></pre id=code>")

	strContent = doCode(strContent, "[sup]", "[/sup]", "<sup>", "</sup>")
	strContent = doCode(strContent, "[sub]", "[/sub]", "<sub>", "</sub>")

	set objRegExp=Nothing
	UBBCode=strContent
end function
Private Function getIP()
    Dim strIPAddr
    If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then
        strIPAddr = Request.ServerVariables("REMOTE_ADDR")
    ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then
        strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1)
    ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then
        strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)
    Else
        strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    End If
    getIP = Trim(Mid(strIPAddr, 1, 30))
End Function
function delbackdb(delfile)
	on error resume next
	dim obj_fso,backdbpath
	backdbpath=server.mappath("movie/"&delfile)
	set obj_fso=server.createobject("scripting.filesystemobject")
	obj_fso.deletefile backdbpath,true
	set obj_fso=nothing
 if err.number<>0 then
		delbackdb=delfile&"上传文件删除失败,只有请管理员从FTP里直接删除了"
	else
		delbackdb="上传文件已从硬盘上被删除,"
	end if
  end function
  function isexists(delfile)
  on error resume next
	dim obj_fso,backdbpath
	backdbpath=server.mappath("movie/"&delfile)
	set obj_fso=server.createobject("scripting.filesystemobject")
     if obj_fso.fileexists(backdbpath) then
	 isexists=true
	 else
	 isexists=false
	 end if
	 end function
	function webpage(webaddr)
	     webpage=true
	    dim left1,right1
		left1=lcase(mid(webaddr,8,1))
		right1=lcase(right(webaddr,1))
		if webaddr="" or left(webaddr,7)<>"http://" then
		      webpage=false
			  end if
	if instr("abcdefghijklmnopqrstuvwxyz1234567890",left1)<=0 then
	       webpage=false
		   end if
	if instr("abcdefghijklmnopqrstuvwxyz1234567890/",right1)<=0 then
	       webpage=false
		   end if
		if instr(webaddr, ".")<=0 then
		    webpage=false
			end if	   
	end function

     Function comp_check(byval str_class)
	on error resume next
	dim obj_check
	set obj_check = server.createobject(str_class)
	set obj_check=nothing
	if err.number<>0 then
		comp_check=false
	else
		comp_check=true
	end if
	err.clear()
    End Function
	function islock(lockip)
	       dim alllock,l
		     islock=true
	alllock=split(lockip , ".")
	 if ubound(alllock)<>3 then
	islock=false
	  else
    for l=0 to 3 
	   if not isnumeric(alllock(l)) then
	        islock=false
			exit for
			end if
	  if alllock(l)>255 or alllock(l)< 0 then
               islock=false
			   exit for
			   end if
			   next
			   end if
          end function
%>
