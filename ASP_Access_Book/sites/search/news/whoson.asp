<!-- #INCLUDE FILE="GrabStats.asp" -->
<%
MyServer=Request.ServerVariables("SERVER_NAME")
MyPath=Request.ServerVariables("SCRIPT_NAME")
MySelf="HTTP://"&MyServer&MyPath
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="REFRESH" CONTENT="20;<%=MySelf%>">
<TITLE>WHOSON</TITLE>
</HEAD>
<BODY>
<%
Application.Lock
localStats=Application("Stats")
Application.UnLock
%>
<CENTER>
<h2>WhosOn</h2>
<TABLE BORDER=1 CELLPADDING=10 
CELLSPACING=0 BGCOLOR="#eeeeee">
<TR BGCOLOR="#cccccc">
    <TH>User ID</TH>
    <TH>Current Page</TH>
    <TH>IP Address</TH>
    <TH>Start Time</TH>
</TR>
<%
FOR i = 0 TO UBOUND( localStats, 2 )
IF localStats( 0, i ) <> "" THEN
%>
<TR>
    <TD><%=localStats( 0, i )%></TD>
    <TD><%=localStats( 1, i )%></TD>
    <TD><%=localStats( 2, i )%></TD>
    <TD><%=localStats( 3, i )%></TD>
</TR>
<%
END IF
NEXT
%>
</TABLE>
</CENTER>
</BODY>
</HTML>
