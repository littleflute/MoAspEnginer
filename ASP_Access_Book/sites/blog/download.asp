<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/ubbcode.asp" -->
<!--#include file="header.asp" -->
<SCRIPT language=JavaScript src="include/common.js" 
type=text/javascript></SCRIPT>
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif" class="wordbreak">
  <tr>
    <td width="160" valign="top" bgcolor="#C8C7BD" nowrap><%
    Call MemberCenter
    Response.Write("<br>")
	Call SiteInfo
    %>
      <br>
      <br>    </td>
<%Dim tb_show
set tb_show=Conn.ExeCute("SELECT tb_Show FROM blog_Toolbox where tb_Name='Download'")
if tb_show("tb_Show") = True Then%>
    <td class="mainbox" background="topBar_bg.gif">
      <%
		Dim CurPage,SQLFiltrate,Url_Add, CateText
	IF trim(Request.QueryString("Page"))<>Empty Then
		Curpage=CLng(CheckStr(Request.QueryString("Page")))
		IF Curpage<0 Then Curpage=1
	Else
		Curpage=1
	End IF
	Dim webLog
	Set webLog=Server.CreateObject("Adodb.Recordset")
if Request.QueryString("action")="icon" then
SQL="SELECT * FROM blog_Download where downl_Cate=1 ORDER BY downl_PostTime desc"
elseif Request.QueryString("action")="wallpaper" then
SQL="SELECT * FROM blog_Download where downl_Cate=2 ORDER BY downl_PostTime desc"
elseif Request.QueryString("action")="other" then
SQL="SELECT * FROM blog_Download where downl_Cate=3 ORDER BY downl_PostTime desc"
else
SQL="SELECT * FROM blog_Download ORDER BY downl_PostTime desc"
end if
	webLog.Open SQL,CONN,1,1
	IF webLog.EOF AND webLog.BOF Then
	Dim tbIcon2
    set tbIcon2=Conn.ExeCute("SELECT tb_Img FROM blog_Toolbox WHERE tb_Name='Download'")
	If tbIcon2("tb_Img")<>empty then 
	Response.Write("<table width='99%' height='52' border='0' align='center' cellpadding='0' cellspacing='6'><tr><td>&nbsp; </td><td width='60%' align='right'><img src='images/download/sd_all.GIF' width='23' height='21' align='absmiddle' border='0'><a href='download.asp'>所有</a>&nbsp;&nbsp;|&nbsp;<img src='images/download/sd_icon.GIF' width='23' height='21' align='absmiddle' border='0'><a href='download.asp?action=icon'>文档</a>&nbsp;&nbsp;|&nbsp;<img src='images/download/sd_wallpaper.GIF' width='23' height='21' align='absmiddle' border='0'>&nbsp;<a href='download.asp?action=wallpaper'>源码</a>&nbsp;&nbsp;|&nbsp;<img src='images/download/sd_other.GIF' align='absmiddle' border='0'>&nbsp;<a href='download.asp?action=other'>软件</a></td></tr></table><table width='99%' border='0' align='center' cellpadding='4' cellspacing='1' class='textbox'><tr><td height='58' align='center'>没有记录</td></tr></table>")
    else
		Response.Write("<table width='99%' height='52' border='0' align='center' cellpadding='0' cellspacing='6'><tr><td>&nbsp; </td><td width='60%' align='right'><img src='images/download/sd_all.GIF' width='23' height='21' align='absmiddle' border='0'><a href='download.asp'>所有</a>&nbsp;&nbsp;|&nbsp;<img src='images/download/sd_icon.GIF' width='23' height='21' align='absmiddle' border='0'><a href='download.asp?action=icon'>文档</a>&nbsp;&nbsp;|&nbsp;<img src='images/download/sd_wallpaper.GIF' width='23' height='21' align='absmiddle' border='0'>&nbsp;<a href='download.asp?action=wallpaper'>源码</a>&nbsp;&nbsp;|&nbsp;<img src='images/download/sd_other.GIF' align='absmiddle' border='0'>&nbsp;<a href='download.asp?action=other'>软件</a></td></tr></table><table width='99%' border='0' align='center' cellpadding='4' cellspacing='1' class='textbox'><tr><td height='58' align='center'>没有记录</td></tr></table>")
    end if
    set tbIcon2=nothing
	Else
		Dim weblog_ID
		webLog.PageSize=4
		webLog.AbsolutePage=CurPage
		Dim Log_Num,MultiPages,PageCount
		Log_Num=webLog.RecordCount
		MultiPages=MultiPage(Log_Num,webLog.PageSize,CurPage,"?")
		if MultiPages<>"" then MultiPages="<span class='multipage'>"&MultiPages&"</span>"
		%>
      <table width='99%' height='52' border='0' align='center' cellpadding='0' cellspacing='0'>
        <tr align="right"> 
          <td colspan="2">  
	<%
	If tbIcon2("tb_Img")<>empty then 
    Response.Write("<img src='"&tbIcon2("tb_Img")&"' width='23' height='21' align='absmiddle'>")
    else
	Response.Write("")
    end if
    set tbIcon2=nothing
    %>

          </td>
        </tr>
        <tr> 
          <td width="40%"> &nbsp;&nbsp; 
            <%If CateText<>"" Then Response.write("<b>"&CateText&"</b> :: ")%> <%=MultiPages%></td>
          <td width="60%" align="right"><img src="images/download/sd_all.GIF" width="23" height="21" align="absmiddle" border="0"><a href="download.asp">所有</a>&nbsp;&nbsp;|&nbsp;<img src="images/download/sd_icon.GIF" width="23" height="21" align="absmiddle" border="0"><a href="download.asp?action=icon">文档</a>&nbsp;&nbsp;|&nbsp;<img src="images/download/sd_wallpaper.GIF" width="23" height="21" align="absmiddle" border="0">&nbsp;<a href="download.asp?action=wallpaper">源码</a>&nbsp;&nbsp;|&nbsp;<img src="images/download/sd_other.GIF" align="absmiddle" border="0">&nbsp;<a href="download.asp?action=other">软件</a></td>
        </tr>
      </table>
	        <%
		Do Until webLog.EOF OR PageCount=webLog.PageSize
			weblog_ID=weblog("downl_ID")
				%>
				<%Dim downl_ID,downl_Cate,downl_Name,downl_Author,downl_From,downl_FromURL,downl_ImgPath,downl_PostTime,downl_Comment,downl_Dcomm1,downl_Dcomm1URL,downl_Dcomm2,downl_Dcomm2URL,downl_Dcomm3,downl_Dcomm3URL,downl_Dcomm4,downl_Dcomm4URL,downl_Nums
		downl_ID=weblog("downl_ID")
		downl_Cate=weblog("downl_Cate")
		downl_Name=weblog("downl_Name")
		downl_Author=weblog("downl_Author")
		downl_From=weblog("downl_From")
		downl_FromURL=weblog("downl_FromURL")
		downl_ImgPath=weblog("downl_ImgPath")
		downl_PostTime=DateToStr(weblog("downl_PostTime"),"YYYY-MM-DD hh:mm:ss","")
		downl_Comment=weblog("downl_Comment")
		downl_Dcomm1=weblog("downl_Dcomm1")
		downl_Dcomm1URL=weblog("downl_Dcomm1URL")
		downl_Dcomm2=weblog("downl_Dcomm2")
		downl_Dcomm2URL=weblog("downl_Dcomm2URL")
		downl_Dcomm3=weblog("downl_Dcomm3")
		downl_Dcomm3URL=weblog("downl_Dcomm3URL")
		downl_Dcomm4=weblog("downl_Dcomm4")
		downl_Dcomm4URL=weblog("downl_Dcomm4URL")
		downl_Nums=weblog("downl_Nums")		
		%>
	  <table width="99%" border="0" align="center" cellpadding="4" cellspacing="1" class="textbox-2">
        <tr> 
          <td colspan="2"  background="images/log_title_bg.jpg" class="textbox-title"><%if downl_Cate=1 then Response.Write("<img src='images/download/d_icon.GIF' width='23' height='21' align='absmiddle'>") Else IF downl_Cate=2 then Response.Write("<img src='images/download/d_wallpaper.GIF' width='23' height='21' align='absmiddle'>")  Else iF downl_Cate=3 then Response.Write("<img src='images/download/d_other.GIF' width='23' height='21' align='absmiddle'>") End if%> &nbsp;<b><%=downl_Name%></b><span class='comment-text'> &nbsp;&nbsp;  
            <%if memStatus=1 then Response.Write("<a href='downledit.asp?downlID="&weblog("downl_ID")&"' title=""编辑此下载"" target=""_self""><img src='images/icon_edit_02.gif' border='0' align='absmiddle'></a>") else Response.Write("") end if%>
            </span></td>
        </tr>
        <tr> 
          <td width="15%" rowspan="2" align="center" valign="top" class="textbox-content"> 
            <% IF weblog("downl_ImgPath")<>Empty Then Response.Write("<img src='"&weblog("downl_ImgPath")&"' width='100' height='100' vspace='4'>") else Response.Write ("<img src='images/nonpreview.gif' width='100' height='100' vspace='4'>") end if %></td>
        </tr>
        <tr> 
          <td height="32" align="right" valign="top" class="textbox-content"> 
            <table width="100%" height="100%" border="0" cellspacing="2" cellpadding="0">
              <tr> 
                <td width="10%">程序发布:</td>
                <td width="18%"><%=downl_Author%></td>
                <td width="25%" align="right" valign="middle"><font color="#999999"><%=downl_Dcomm1%></font></td>
                <td width="1%" align="right" valign="middle"><a href="down.asp?downID=<%=weblog("downl_ID")%>&action=Url_1" target="_blank"><img src="images/downloadbm.gif" align="absmiddle" border="0"></a></td>
              </tr>
              <tr> 
                <td>发布日期:</td>
                <td><%=DateToStr(webLog("downl_PostTime"),"Y-m-d H:I")%></td>
                <td align="right"><%if downl_Dcomm2<>empty then Response.Write("<font color='#999999'>"&weblog("downl_Dcomm2")&"</font>") else Response.Write("<font color='#999999'>无</font>") end if%></td>
                <td align="right"><%if downl_Dcomm2URL<>empty then Response.Write("<a href='down.asp?downID="&weblog("downl_ID")&"&action=Url_2' target='_blank'><img src='images/downloadbm.gif' align='absmiddle' border='0'></a>") else Response.Write("<img src='images/nondownload.gif' align='absmiddle' border='0'>") end if%></td>
              </tr>
              <tr> 
                <td>程序大小:</td>
                <td><%=weblog("downl_size")%></td>
                <td align="right"><%if downl_Dcomm3<>empty then Response.Write("<font color='#999999'>"&weblog("downl_Dcomm3")&"</font>") else Response.Write("<font color='#999999'>无</font>") end if%></td>
                <td align="right"><%if downl_Dcomm3URL<>empty then Response.Write("<a href='down.asp?downID="&weblog("downl_ID")&"&action=Url_3' target='_blank'><img src='images/downloadbm.gif' align='absmiddle' border='0'></a>") else Response.Write("<img src='images/nondownload.gif' align='absmiddle' border='0'>") end if%></td>
              </tr>
              <tr> 
                <td>下载次数:</td>
                <td><%=downl_Nums%>&nbsp;次</td>
                <td align="right"><%if downl_Dcomm4<>empty then Response.Write("<font color='#999999'>"&weblog("downl_Dcomm4")&"</font>") else Response.Write("<font color='#999999'>无</font>") end if%></td>
                <td align="right"><%if downl_Dcomm4URL<>empty then Response.Write("<a href='down.asp?downID="&weblog("downl_ID")&"&action=Url_4' target='_blank'><img src='images/downloadbm.gif' align='absmiddle' border='0'></a>") else Response.Write("<img src='images/nondownload.gif' align='absmiddle' border='0'>") end if%></td>
              </tr>
              <tr> 
                <td>程序介绍:</td>
                <td colspan="2" align="left"><SPAN style="CURSOR: hand" onclick="showIntro('<%=weblog("downl_ID")%>');">点击查看详细介绍</SPAN></td>
                <td></td>
              </tr>
            </table></td>
        </tr>
      <td colspan="4" id=<%=weblog("downl_ID")%> style="DISPLAY: none"><%=HtmlEncode(webLog("downl_Comment"))%></td>
        </tr>
      </table>
      <br>
	   <%
			webLog.MoveNext
			PageCount=PageCount+1
		Loop
						End If
	webLog.Close
	Set webLog=Nothing%>
	  <table width='99%' height='5' border='0' align='center' cellpadding='0' cellspacing='6'>
      <tr>
      <td><%=MultiPages%></td>
      </tr>
      </table>
</table>
</td>
  </tr>
</table>
<%else%>
    <td valign="top"><br><br><table border="0"  cellpadding="4" cellspacing="1" class="textbox"><tr>
      <td align="center"><br>
        对不起,此功能被关闭。<br>
        <br></td></tr></table></td>
	</tr>
</table>
<%set tb_show=nothing
end if%><!--#include file="footer.asp" -->