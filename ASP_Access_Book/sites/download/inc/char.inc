<%
function htmlencode1(fString)
if fString<>"" and not isnull(fString) then
fString = replace(fString, "&gt;", ">")
fString = replace(fString, "&lt;", "<")

fString = Replace(fString, "&nbsp;", chr(32))
fString = Replace(fString, "</P><P>", CHR(10) & CHR(10))
fString = Replace(fString, "<BR>", CHR(10))
htmlencode1=fString
else
htmlencode1="1"
end if
end function

function htmlencode2(fString)
if fString<>"" and not isnull(fString) then

fString = replace(fString, ">", "&gt;")
fString = replace(fString, "<", "&lt;")

fString = Replace(fString, chr(32), "&nbsp;")
fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
fString = Replace(fString, CHR(10), "<BR>")
htmlencode2=fString
else
htmlencode2="1"
end if
end function
function htmlencode3(fString)
if fString<>"" and not isnull(fString) then

fString = replace(fString, "&gt;", ">")
fString = replace(fString, "&lt;", "<")

fString = Replace(fString, "&nbsp;", chr(32))
fString = Replace(fString, "</P><P>", CHR(10) & CHR(10))
fString = Replace(fString, "<BR>", CHR(10))
htmlencode3=fString
else
htmlencode3="1"
end if
end function

function htmlContent(fString)
if fString<>"" and not isnull(fString) then
fString = replace(fString, "&gt;", ">")
fString = replace(fString, "&lt;", "<")
fString = Replace(fString, "&nbsp;", chr(32))
fString = Replace(fString, "</P><P>", CHR(10))
fString = Replace(fString, "<BR>", CHR(10))
fString = replace(fString,"<p>","")
fString = replace(fString,"<p","<")
fString = replace(fString,"</P>",CHR(10))
fString = replace(fString,"[[","<")
fString = replace(fString,"]]",">")
htmlContent=fString
else
htmlContent="1"
end if
end function
%>