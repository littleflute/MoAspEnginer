<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/library.asp" -->
<!--#include file="include/ubbcode.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <table width="768" border="0" align="center" cellpadding="4" cellspacing="0" class="wordbreak">
      <tr>
        <td valign="top" nowrap bgcolor="#FFFFFF" class="bk_1"><div class="siderbar_head">最新日志列表</div>
            <div class="content_main">
              <%Dim RS
                Set RS=Server.CreateObject("Adodb.RecordSet")
                SQL="SELECT TOP 10 log_ID,log_Title,log_Weather FROM blog_Content WHERE log_IsShow=True ORDER BY log_PostTime DESC"
                RS.Open SQL,Conn,1,1
                SQLQueryNums=SQLQueryNums+1
                If RS.Eof And RS.Bof Then
                    Response.Write("None")
                Else
                    Dim log_Weather,log_ID
                    Do While Not RS.Eof
                        log_ID=RS("log_ID")
                        log_Weather=Split(RS("log_Weather"),"|")
                        Response.Write("<div class=""newline"" style=""width:294px;""><a href=""blogview.asp?logID="&log_ID&""">"&HTMLEncode(cutStr(RS("log_Title"),20))&"</a></div><table width=""100%"" height=""9"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td></td></tr></table>")
                        RS.Movenext
                    Loop
                End If
                RS.Close
                Set RS=Nothing%>
            </div>
        <td colspan="2" valign="top" nowrap bgcolor="#FFFFFF" class="bk_1"><div class="siderbar_head">最新主题列表</div>
            <div class="content_main">
              <%
                Set RS=Server.CreateObject("Adodb.RecordSet")
                SQL="SELECT TOP 10 thread_ID,thread_forumID,thread_Title,thread_Icon FROM blog_Threads WHERE thread_IsClose=False ORDER BY thread_PostTime DESC"
                RS.Open SQL,Conn,1,1
                SQLQueryNums=SQLQueryNums+1
                If RS.Eof And RS.Bof Then
                    Response.Write("None")
                Else
                    Dim thread_Icon
                    Do While Not RS.Eof
                        thread_Icon=RS("thread_Icon")
                        If thread_Icon=Empty Then thread_Icon="../newwin.gif"
                        Response.Write("<div class=""newline"" style=""width:294px;""><img src=""images/threadicon/"&thread_Icon&""" alt="""" align=""absmiddle"">&nbsp;<a href=""threadview.asp?forumID="&RS("thread_forumID")&"&threadID="&RS("thread_ID")&""">"&HTMLEncode(cutStr(RS("thread_Title"),40))&"</a></div><table width=""100%"" height=""9"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td></td></tr></table>")
                        RS.Movenext
                    Loop
                End If
                RS.Close
                Set RS=Nothing%>
            </div>
            <div class="siderbar_head"> </div>
      </tr>
      <tbody>
        <tr>
          <td width="312" valign="top" style="padding-right:5px;"><div class="siderbar_head">最热日志列表</div>
              <div class="content_main">
                <%
                Set RS=Server.CreateObject("Adodb.RecordSet")
                SQL="SELECT TOP 10 log_ID,log_Title,log_Weather FROM blog_Content WHERE log_IsShow=True ORDER BY log_ViewNums DESC"
                RS.Open SQL,Conn,1,1
                SQLQueryNums=SQLQueryNums+1
                If RS.Eof And RS.Bof Then
                    Response.Write("None")
                Else
                    Do While Not RS.Eof
                        log_ID=RS("log_ID")
                        log_Weather=Split(RS("log_Weather"),"|")
                        Response.Write("<div class=""newline"" style=""width:294px;""><img src=""images/weather/"&log_Weather(0)&".gif"" alt="""&log_Weather(1)&""" align=""absmiddle"">&nbsp;<a href=""blogview.asp?logID="&log_ID&""">"&HTMLEncode(cutStr(RS("log_Title"),40))&"</a></div><table width=""100%"" height=""9"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td></td></tr></table>")
                        RS.Movenext
                    Loop
                End If
                RS.Close
                Set RS=Nothing%>
            </div></td>
          <td width="440" valign="top" style="padding-left:5px;"><div class="siderbar_head">最热主题列表</div>
              <div class="content_main">
                <%
                Set RS=Server.CreateObject("Adodb.RecordSet")
                SQL="SELECT TOP 5 thread_ID,thread_forumID,thread_Title,thread_Icon FROM blog_Threads WHERE thread_IsClose=False ORDER BY thread_ViewNums DESC"
                RS.Open SQL,Conn,1,1
                SQLQueryNums=SQLQueryNums+1
                If RS.Eof And RS.Bof Then
                    Response.Write("None")
                Else
                    Do While Not RS.Eof
                        thread_Icon=RS("thread_Icon")
                        If thread_Icon=Empty Then thread_Icon="../newwin.gif"
                        Response.Write("<div class=""newline"" style=""width:294px;""><img src=""images/threadicon/"&thread_Icon&""" alt="""" align=""absmiddle"">&nbsp;<a href=""threadview.asp?forumID="&RS("thread_forumID")&"&threadID="&RS("thread_ID")&""">"&HTMLEncode(cutStr(RS("thread_Title"),40))&"</a></div><table width=""100%"" height=""9"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td></td></tr></table>")
                        RS.Movenext
                    Loop
                End If
                RS.Close
                Set RS=Nothing%>
            </div></td>
        </tr>
      </tbody>
      <tbody>
      </tbody>
    </table>
    <td valign="top" bgcolor="#FFFFFF">    
    