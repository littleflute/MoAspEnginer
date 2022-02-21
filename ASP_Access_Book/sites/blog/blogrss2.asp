<!--#include file="commond.asp" -->
<%
Response.ContentType="application/xml"
Response.Expires=0
Response.Write("<?xml version=""1.0"" encoding=""UTF-8""?>")
%>
<!--#include file="include/function.asp" -->
<!--#include file="include/ubbcode.asp" -->
<rss>
<channel>
<title><%=SiteName%></title>
<link><%=SiteURL%></link>
<Description><%=SiteName%></Description>
<language>zh-cn</language>
<copyright>Copyright 2003-2004 Loveyuki</copyright>
<webMaster>webmaster@loveyuki.com</webMaster>
<image>
	<title><%=SiteName%></title> 
	<url><%=SiteURL%>/images/logos.gif</url> 
	<link><%=SiteURL%></link> 
	<description><%=SiteName%></description> 
</image>
<%
Dim cate_ID
cate_ID=Request.QueryString("cateID")
IF IsInteger(cate_ID) = False Then
	SQL="SELECT TOP 10 L.log_ID,L.log_Title,l.log_Author,L.log_PostTime,L.log_Content,L.log_IsShow,L.log_DisSM,L.log_DisUBB,L.log_DisIMG,L.log_AutoURL,L.log_AutoKEY,C.cate_Name FROM blog_Content AS L,blog_Category AS C WHERE log_IsShow=True AND C.cate_ID=L.log_cateID ORDER BY log_PostTime DESC"
Else
	SQL="SELECT TOP 10 L.log_ID,L.log_Title,l.log_Author,L.log_PostTime,L.log_Content,L.log_IsShow,L.log_DisSM,L.log_DisUBB,L.log_DisIMG,L.log_AutoURL,L.log_AutoKEY,C.cate_Name FROM blog_Content AS L,blog_Category AS C WHERE log_cateID="&cate_ID&" AND log_IsShow=True AND C.cate_ID=L.log_cateID ORDER BY log_PostTime DESC"
End IF
Dim RS
Set RS=Server.CreateObject("ADODB.RecordSet")
RS.Open SQL,Conn,1,1
IF RS.EOF AND RS.BOF Then
	Response.Write("<item></item>")
Else
	Do While NOT RS.EOF
		Response.Write("<item>")
		Response.Write("<link>"&SiteURL&"/blogview.asp?logID="&RS("log_ID")&"</link>")
		Response.Write("<title><![CDATA["&RS("log_Title")&"]]></title>")
		Response.Write("<author>"&RS("log_Author")&"</author>")
		Response.Write("<category>"&RS("cate_Name")&"</category>")
		Response.Write("<pubDate>"&RS("log_PostTime")&"</pubDate>")
		Response.Write("<guid>"&SiteURL&"/blogview.asp?logID="&RS("log_ID")&"</guid>")
		Response.Write("<description><![CDATA["&Ubbcode(HTMLEncode(RS("log_Content")),RS("log_DisSM"),RS("log_DisUBB"),RS("log_DisIMG"),RS("log_AutoURL"),RS("log_AutoKEY"))&"]]></description>")
		Response.Write("</item>")
		RS.MoveNext
	Loop
End IF
RS.Close
Set RS=Nothing
Conn.Close
Set Conn=Nothing
%>
</channel>
</rss>