<%
Dim ServerObject(9)
ServerObject(9) = "Scripting.FileSystemObject"

Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function

'�ȵ�ͼƬ
function HotImg(NewsID,i)
If  not IsObjInstalled(ServerObject(9)) Then
Response.Write "<img src='"&ImgPath& NewsID & "-" & i & ".jpg' border=0 width='120' "&imgheight&" alt=""��֧�� FSO! ֻ����ʾjpgͼƬ"">"
else
On Error Resume Next	
set DelectFile=server.CreateObject("scripting.filesystemobject")
CurrentPath=server.MapPath(ImgPath)+"/"
FileName=CurrentPath   & NewsID & "-" & i & ".jpg"
if DelectFile.FileExists(FileName) then
HotImg="<img src='"&ImgPath& NewsID & "-" & i & ".jpg' border=0 width='120' "&imgheight&">"
exit function
else
FileName=CurrentPath   & NewsID & "-" & i & ".gif"
if DelectFile.FileExists(FileName) then
HotImg="<img src='"&ImgPath& NewsID & "-" & i & ".gif' border=0 width='120' "&imgheight&">"
exit function
else
FileName=CurrentPath   & NewsID & "-" & i & ".png"
if DelectFile.FileExists(FileName) then
HotImg="<img src='"&ImgPath& NewsID & "-" & i & ".png' border=0 width='120' "&imgheight&">"
exit function
else
FileName=CurrentPath   & NewsID & "-" & i & ".swf"
if DelectFile.FileExists(FileName) then
HotImg="<object Classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='120' "&imgheight&"><param name=movie value='"&ImgPath& NewsID & "-" & i & ".swf'><param name=quality value=high><embed src='"&ImgPath& NewsID & "-" & i & ".swf' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='120'></embed></object>"
exit function
else
HotImg="<img src='"&ImgPath& NewsID & "-" & i & ".bmp' border=0 width='120' "&imgheight&">"
exit function
end if
end if
end if
end if
end if
end function

'���ͼƬ
function DelectImageFile(NewsID,i)
If  not IsObjInstalled(ServerObject(9)) Then
Response.Write "<img src='"&ImgPath& NewsID & "-" & i & ".jpg' height=200 border=0 alt=""��֧�� FSO! ֻ����ʾjpgͼƬ"">"
else	
set DelectFile=server.CreateObject("scripting.filesystemobject")
CurrentPath=server.MapPath(ImgPath)+"/"
FileName=CurrentPath   & NewsID & "-" & i & ".jpg"
if DelectFile.FileExists(FileName) then
DelectImageFile="<img src='"&ImgPath& NewsID & "-" & i & ".jpg' border=0>"
exit function
else
FileName=CurrentPath   & NewsID & "-" & i & ".gif"
if DelectFile.FileExists(FileName) then
DelectImageFile="<img src='"&ImgPath& NewsID & "-" & i & ".gif' border=0>"
exit function
else
FileName=CurrentPath   & NewsID & "-" & i & ".png"
if DelectFile.FileExists(FileName) then
DelectImageFile="<img src='"&ImgPath& NewsID & "-" & i & ".png' border=0>"
exit function
else
FileName=CurrentPath   & NewsID & "-" & i & ".swf"
if DelectFile.FileExists(FileName) then
DelectImageFile="<object Classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width=300><param name=movie value='"&ImgPath& NewsID & "-" & i & ".swf'><param name=quality value=high><embed src='"&ImgPath& NewsID & "-" & i & ".swf' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width="&FlashWidth&"></embed></object>"
exit function
else
DelectImageFile="<img src='"&ImgPath& NewsID & "-" & i & ".bmp' border=0>"
exit function
end if
end if
end if
end if
end if
end function

'�ϴ�ͼƬ
function DelectImageFile_Upload(NewsID,i)
set DelectFile=server.CreateObject("scripting.filesystemobject")
CurrentPath=server.MapPath("&ImgPath&")
FileName=CurrentPath   & NewsID & "-" & i & ".gif"
if DelectFile.FileExists(FileName) then
DelectImageFile_Upload= NewsID & "-" & i & ".gif"
exit function
else
FileName=CurrentPath   & NewsID & "-" & i & ".jpg"
if DelectFile.FileExists(FileName) then
DelectImageFile_Upload= NewsID & "-" & i & ".jpg"
exit function
else
FileName=CurrentPath   & NewsID & "-" & i & ".png"
if DelectFile.FileExists(FileName) then
DelectImageFile_Upload= NewsID & "-" & i & ".png"
exit function
else
FileName=CurrentPath   & NewsID & "-" & i & ".swf"
if DelectFile.FileExists(FileName) then
DelectImageFile_Upload= NewsID & "-" & i & ".swf"
exit function
else
FileName=CurrentPath   & NewsID & "-" & i & ".bmp"
if DelectFile.FileExists(FileName) then
DelectImageFile_Upload= NewsID & "-" & i & ".bmp"
exit function
else
DelectImageFile_Upload=""
exit function
end if
end if
end if
end if
end if
end function

'����ͼƬ���ô���
Function HtmlSelfEnCode(content,ImageNum)
Image=ImageNum
TempContent=content
if image>0 then
for i=1 to image
TempContent=replace(TempContent,"[[image" & i & "]]","" & DelectImageFile(NewsID,i) & "")
next
end if
TempContent=replace(TempContent,"[[left]]","<table border=0 cellspacing=5 cellpadding=0 align=left><tr><td>")
TempContent=replace(TempContent,"[[/left]]","</td></tr></table>")
TempContent=replace(TempContent,"[[center]]","<table border=0 cellspacing=5 cellpadding=0 align=center><tr><td>")
TempContent=replace(TempContent,"[[/center]]","</td></tr></table>")
TempContent=replace(TempContent,"[[right]]","<table border=0 cellspacing=5 cellpadding=0 align=right><tr><td>")
TempContent=replace(TempContent,"[[/right]]","</td></tr></table>")
TempContent=replace(TempContent,"[[","<")
TempContent=replace(TempContent,"]]",">")
HtmlSelfEnCode=TempContent
End Function
	
function checkOverFlow(strChinese, lenMaxWord)
'�ж��ַ������Ƿ����
'strChinese Ϊ������ַ���,lenMaxWord Ϊ���Ƶ��ַ�����

dim i, lenTotal, strWord , firstChinese

if strChinese = "" or vartype(strChinese) = vbNull or CLng(lenMaxWord) <= 0 then
checkOverFlow = False
exit function
end if

lenTotal = 0
for i=1 to Len(strChinese)
strWord = mid(strChinese, i, 1)
if asc(strWord) < 0 or asc(strWord) > 127 then
lenTotal = lenTotal + 2
else
lenTotal = lenTotal + 1
end if
next

'�ж��ַ��Ƿ����
if lenTotal > lenMaxWord then
checkOverFlow = True
else
checkOverFlow = False
end if
end function

function GetTrueLength(strChinese, lenMaxWord, strSpaceBar)
'��ȡ��ȷ��Ӣ��/���ֳ���
'strChinese Ϊ������ַ���,lenMaxWord Ϊ���Ƶ��ַ�����

dim i, j, strTail, lenTotal, lenWord
dim strWord, bOverFlow, RetString

if strChinese = "" or vartype(strChinese) = vbNull or CLng(lenMaxWord) <= 0 then
GetTrueLength = ""
exit function
end if

strTail = "��"
bOverFlow = False

lenTotal = 0
for i=1 to Len(strChinese)
strWord = mid(strChinese, i, 1)
if asc(strWord) < 0 or asc(strWord) > 127 then
lenTotal = lenTotal + 2
else
lenTotal = lenTotal + 1
end if
next
'�ж��ַ��Ƿ����
if lenTotal > lenMaxWord then bOverFlow = True

strSpaceBar = ""
if bOverFlow = True then
'�ַ����,ȥβ
lenWord = 0
RetString = ""
for i=1 to Len(strChinese)
strWord = mid(strChinese, i, 1)
if asc(strWord) < 0 or asc(strWord) > 127 then lenNow = 2 else lenNow = 1
lenWord = lenWord + lenNow
'�ص����ಿ��
if lenWord <= (lenMaxWord - Len(strTail)) then
RetString = RetString + strWord
else
RetString = RetString + strTail
lenWord = lenWord + Len(strTail) - lenNow
if (lenMaxWord-lenWord)>0 then
for j =1 to lenMaxWord-lenWord
strSpaceBar = strSpaceBar + "&nbsp;"
next
end if
GetTrueLength = RetString
exit for
end if
next
else
'�ַ������,����λ
RetString = strChinese
if (lenMaxWord-lenTotal)>0 then
for i =1 to lenMaxWord-lenTotal
strSpaceBar = strSpaceBar + "&nbsp;"
next
end if
GetTrueLength = RetString
end if
end function

'��������ͨ��ѡ���
NoContent=" NewsID,Title,model,BigClassName,SmallClassName,SpecialName,author,original,UpdateTime,image,click,goodnews "

function NewsUrl	'�������ű���URL
model=rs("model")
if model=0 then
model=""
end if
newsurl="shownews"&model&".asp?newsid=" & rs("NewsID") 	
end function

function showTitle(strClass,strMaxLen)	'������⼰����
'strClass Ϊ��ʾ��ʽ����class="��ʽ"��ֵ��������˫���ű�ʾ��
'strMaxLen Ϊ��ʾ���ȣ�ż����
strSubject = HTMLDecode(rs("Title"))	
strTrueSubject = GetTrueLength(strSubject, strMaxLen, strSpaceBar)
m_bOverFlow = checkOverFlow(strSubject, strMaxLen)
if m_bOverFlow = True then
strTip = strSubject
else
strTip = ""
end if
if strClass="" then strClass="MainContentS"
Response.Write "<a class='"&strClass&"' href='"&newsurl&"' title='"&strTip&"' target='_blank'>"&strTrueSubject&"</a>"
end function			
			
function showTime      		'����ʱ����ʾ��ʽ
'����Ĭ�ϵ�ΪNEWʱ������Ϊ��ɫ
if DateValue(rs("updatetime"))=>DateValue(date()-Indate) then
fontcolor="<font color='"&AlertFColor&"'>"
else
fontcolor="<font Class=TitleMore>"
end if
Response.Write "&nbsp;<font Class=TitleMore>(" & fontcolor & DateValue(rs("UpdateTime"))&"</font>)</font>"
end function

function showImg					'������ͼ�����ű�־
if rs("image")>0 then showImg="<font Class=TitleMore>[<font color='"&AlertFColor&"'>ͼ</font>]</font>"
end function

function showClick					'��������ʽ
showClick="<font Class=TitleMore>[<font color='"&AlertFColor&"'>" & rs("click") &"</font>]</font>"	
end function

function HTMLDecode(fString)
fString = replace(fString, "&amp;", "&")
fString = replace(fString, "&gt;", ">")
fString = replace(fString, "&lt;", "<")
fString = replace(fString, "&quot;", Chr(34))
fString = Replace(fString, "��", "...")
HTMLDecode = fString
end function

Function Space(strHeight)					'����������֮��ļ��
if strHeight="" then strHeight=THeight
if strHeight<>0 then Response.Write "<table width='90%' border='0' height='"&strHeight&"' cellpadding=""0"" cellspacing=""0""><tr><td></td></tr></table>"
end function

Function trline()					'�����й�ҳ��Ŀ�����֮��ķָ���
Response.Write "<tr><td width=""100%"" bgcolor=""#e4e4e4"" height=""6""></td></tr><tr><td height=""6""></td></tr>"
end function

Function OutTable(strside)					'��������ʽ
if strside="left" then Response.Write "<TD BGCOLOR=""#ffffff"" WIDTH=""1""></TD><TD BGCOLOR='"&Top4bgColor&"' WIDTH=""7""></TD><TD BGCOLOR=""#000000"" WIDTH=""1""></TD>"	
if strside="right" then Response.Write "<TD BGCOLOR=""#000000"" WIDTH=""1""></TD><TD BGCOLOR='"&Top4bgColor&"' WIDTH=""7""></TD><TD BGCOLOR=""#ffffff"" WIDTH=""1""></TD>"
end function

Function InTable(strside)					'�����ڿ��ʽ
if strside="left" then Response.Write "<TD BGCOLOR=""#666666"" WIDTH=""1""></TD>"	'��������
if strside="right" then Response.Write "<TD background=""images/lline.gif"" WIDTH=""1""></TD>"  '��������
if strside="bottoml" then Response.Write "<tr><TD BGCOLOR="&LeftBColor&" HEIGHT='1'></TD></tr>"   '������
if strside="bottomr" then Response.Write "<tr><TD BGCOLOR="&RightBColor&" HEIGHT='1'></TD></tr>"	'�Һ����
if strside="middle1" then Response.Write "<tr><TD background=""images/hline.gif"" HEIGHT=""1""></TD></tr>"	'�к�����޷���
if strside="middle2" then Response.Write "<tr><TD background=""images/hline.gif"" HEIGHT=""1"" colspan=""2""></TD></tr>"	'�к�����з���
end function		
%>