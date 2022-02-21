<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/md5code.asp" -->
<!--#include file="header.asp" -->
<%IF memStatus<>"Admin" AND memStatus<>"SupAdmin" Then%>
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
    <td width="128" align="center" nowrap><b>发表新日志<br>
          <br>
    <font color="#FF0000">出现错误</font> </b></td>
    <td align="center" valign="top"><br>
      <br><div class="msg_head">没有权限发表新日志</div>
        <div class="msg_content"><a href="default.asp">点击返回主页面</a><br><br>
              <a href="logging.asp">或者用管理员帐户重新登陆</a></div>
    <br>
    <br></td>
  </tr>
</table>
<%Else
IF Request.Form("log_CateID")=Empty Then%>
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
    <td width="128" align="center" nowrap><b>发表新日志<br>
      <br>
      <font color="#FF0000">选择分类</font>
    </b></td>
    <td align="center" valign="top"><table width="90%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
      <tr align="center">
        <td colspan="2" class="msg_head">选择分类</td>
        </tr>
      <form name="form_C" method="post" action=""><tr bgcolor="#FFFFFF">
        <td width="19%" align="right"><b>分类：</b></td>
        <td width="81%"><select name="log_CateID" id="log_CateID">
          <option value="">请选择分类</option>
		  <%Dim log_CateNums,log_CateNumI
			log_CateNums=Ubound(Arr_Category,2)
			For log_CateNumI=0 To log_CateNums
				Response.Write("<option value='"&Arr_Category(0,log_CateNumI)&"'>"&Arr_Category(1,log_CateNumI)&"</option>")
			Next%>
        </select>
        </td>
      </tr>
      <tr align="center" bgcolor="#FFFFFF">
        <td colspan="2"><input name="C_Submit" type="submit" id="C_Submit" value=" 确 定 "></td>
        </tr></form>
    </table></td>
  </tr>
</table><%ElseIF Request.Form("PostBLOG")=Empty Then%>
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
	<td width="100%" valign="top"><table width="97%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
        <tr align="center">
          <td colspan="2" class="msg_head">在分类&nbsp;<font color="#FFFA20">[ <%=Conn.ExeCute("SELECT cate_Name FROM blog_Category WHERE cate_ID="&Request.Form("log_CateID")&"")(0)%> ]</font>&nbsp;中发表新日志</td>
        </tr>
        <form name="input" method="post" action="">
          <tr bgcolor="#FFFFFF">
            <td width="112" align="right" nowrap><b>标题：</b></td>
            <td width="100%">
              <input name="log_Title" type="text" id="log_Title" size="50">&nbsp;<select name="log_Weather" id="log_Weather" onChange="document.images['show_Weather'].src='images/weather/'+options[selectedIndex].value.split('|')[0]+'.gif';" size="1">
                    <option value="0|未知" selected>天气</option>
                    <option value="1|晴天">晴天</option>
                    <option value="2|多云">多云</option>	 
                    <option value="3|雨天">雨天</option>
                    <option value="4|刮风">刮风</option>
                    <option value="5|雪天">雪天</option>
                    <option value="6|彩虹">彩虹</option>
                    <option value="7|露水">露水</option>
                  </select>&nbsp;<img id="show_Weather" id="show_Weather" src="images/weather/0.gif" align="absmiddle">
			</td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td align="right"><strong>属性：</strong></td>
            <td><input name="log_IsShow" type="radio" value="0" checked> 公开日志&nbsp;&nbsp;<input type="radio" name="log_IsShow" value="1"> 隐藏日志&nbsp;&nbsp;|&nbsp;&nbsp;<input name="log_IsTop" type="checkbox" id="log_IsTop" value="1"> 置顶日志&nbsp;&nbsp;<input name="log_DisComment" type="checkbox" id="log_DisComment" value="1"> 禁止评论</td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td align="right"><b>来自：</b></td>
            <td><input name="log_From" type="text" id="log_From" value="本站原创" size="12">
&nbsp;<b>地址：
              <input name="log_FromURL" type="text" id="log_FromURL" value="<%=siteURL%>" size="38">
            </b></td>
          </tr>
<tr bgcolor="#FFFFFF">
            <td align="right"><b>Tags：</b></td>
            <td><input name="log_Tags" type="text" id="log_Tags" size="50"> 最多5个Tag 用;分隔</td>
          </tr>
<script language="JavaScript" src="include/ubbhelp.js"></script>
<script language="JavaScript" src="include/ubbcode.js"></script>
          <tr bgcolor="#FFFFFF">
            <td align="right" valign="top"><b>内容：<br>
              <br>
            </b>
              <table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
                <tr>
                  <td align="center" bgcolor="#EFEFEF"><b>表&nbsp;&nbsp;情</b></td>
                </tr>
                <tr>
                  <td bgcolor="#FFFFFF">
			<%Dim log_SmiliesListNums,log_SmiliesListNumI
			log_SmiliesListNums=Ubound(Arr_Smilies,2)
			TempVar=""
			For log_SmiliesListNumI=0 To log_SmiliesListNums
				Response.Write(TempVar&"<img style=""cursor:hand;"" onclick=""AddText('"&Arr_Smilies(2,log_SmiliesListNumI)&"');"" src=""images/smilies/"&Arr_Smilies(1,log_SmiliesListNumI)&""" />")
				TempVar=" "
			Next%></td>
                </tr>
              </table>
              <div style="padding-left:8px;" align="left" width="100%">
                <br>
                <br>
                <input name="log_DisSM" type="checkbox" id="log_DisSM" value="1">
                禁止表情<br>
                <input name="log_DisUBB" type="checkbox" id="log_DisUBB" value="1">
                禁止UBB<br>
                <input name="log_DisIMG" type="checkbox" id="log_DisIMG" value="1">
                禁止图片<br>
				<input name="log_AutoURL" type="checkbox" id="log_AutoURL" value="1" checked>
                识别链接<br>
				<input name="log_AutoKEY" type="checkbox" id="log_AutoKEY" value="1">
                识别关键字
              </div></td>
            <td><table width="98%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td width="100%"><select name="font" onFocus="this.selectedIndex=0" onChange="chfont(this.options[this.selectedIndex].value)" size="1">
                    <option value="" selected>选择字体</option>
                    <option value="宋体">宋体</option>
                    <option value="黑体">黑体</option>
                    <option value="Arial">Arial</option>
                    <option value="Book Antiqua">Book Antiqua</option>
                    <option value="Century Gothic">Century Gothic</option>
                    <option value="Courier New">Courier New</option>
                    <option value="Georgia">Georgia</option>
                    <option value="Impact">Impact</option>
                    <option value="Tahoma">Tahoma</option>
                    <option value="Times New Roman">Times New Roman</option>
                    <option value="Verdana">Verdana</option>
                  </select>
                    <select name="size" onFocus="this.selectedIndex=0" onChange="chsize(this.options[this.selectedIndex].value)" size="1">
                      <option value="" selected>字体大小</option>
                      <option value="-2">-2</option>
                      <option value="-1">-1</option>
                      <option value="1">1</option>
                      <option value="2">2</option>
                      <option value="3">3</option>
                      <option value="4">4</option>
                      <option value="5">5</option>
                      <option value="6">6</option>
                      <option value="7">7</option>
                    </select>
                    <select name="color"  onFocus="this.selectedIndex=0" onChange="chcolor(this.options[this.selectedIndex].value)" size="1">
                      <option value="" selected>字体颜色</option>
                      <option value="White" style="background-color:white;color:white;">White</option>
                      <option value="Black" style="background-color:black;color:black;">Black</option>
                      <option value="Red" style="background-color:red;color:red;">Red</option>
                      <option value="Yellow" style="background-color:yellow;color:yellow;">Yellow</option>
                      <option value="Pink" style="background-color:pink;color:pink;">Pink</option>
                      <option value="Green" style="background-color:green;color:green;">Green</option>
                      <option value="Orange" style="background-color:orange;color:orange;">Orange</option>
                      <option value="Purple" style="background-color:purple;color:purple;">Purple</option>
                      <option value="Blue" style="background-color:blue;color:blue;">Blue</option>
                      <option value="Beige" style="background-color:beige;color:beige;">Beige</option>
                      <option value="Brown" style="background-color:brown;color:brown;">Brown</option>
                      <option value="Teal" style="background-color:teal;color:teal;">Teal</option>
                      <option value="Navy" style="background-color:navy;color:navy;">Navy</option>
                      <option value="Maroon" style="background-color:maroon;color:maroon;">Maroon</option>
                      <option value="LimeGreen" style="background-color:limegreen;color:limegreen;">LimeGreen</option>
                  </select><b>模式：</b>&nbsp;
                    <input type="radio" name="mode" value="2" onclick="chmode('2')" checked>
      基本&nbsp;
      <input type="radio" name="mode" value="0" onclick="chmode('0')">
      高级&nbsp;
      <input type="radio" name="mode" value="1" onclick="chmode('1')">
      帮助</td>
                <td rowspan="2" nowrap></td>
              </tr>
               <tr>
                <td><a href="javascript:bold()"><img src="images/ubbcode/bb_bold.gif" border="0" alt="插入粗体文本"></a> <a href="javascript:italicize()"><img src="images/ubbcode/bb_italicize.gif" border="0" alt="插入斜体文本"></a> <a href="javascript:underline()"><img src="images/ubbcode/bb_underline.gif" border="0" alt="插入下划线"></a> <a href="javascript:center()"><img src="images/ubbcode/bb_center.gif" border="0" alt="居中对齐"></a> <a href="#ubb" onclick="javascript:sub()"><img src="images/ubbcode/bb_sub.gif" border="0" alt="文字下标"></a> <a href="#ubb" onclick="javascript:sup()"><img src="images/ubbcode/bb_sup.gif" border="0" alt="文字上标"></a> <a href="javascript:hyperlink()"><img src="images/ubbcode/bb_url.gif" border="0" alt="插入超级链接"></a> <a href="javascript:email()"><img src="images/ubbcode/bb_email.gif" border="0" alt="插入邮件地址"></a> <a href="javascript:image()"><img src="images/ubbcode/bb_image.gif" border="0" alt="插入图像"></a> <a href="javascript:code()"><img src="images/ubbcode/bb_code.gif" border="0" alt="插入代码"></a> <a href="javascript:quote()"><img src="images/ubbcode/bb_quote.gif" border="0" alt="插入引用"></a> <a href="javascript:list()"><img src="images/ubbcode/bb_list.gif" border="0" alt="插入列表"></a> <a href="javascript:flash()"><img src="images/ubbcode/bb_flash.gif" border="0" alt="插入 Flash"></a> <a href="javascript:wma()"><img src="images/ubbcode/bb_wma.gif" alt="插入音频文件" width="23" height="22" border="0"></a> <a href="javascript:wmv()"><img src="images/ubbcode/bb_wmv.gif" alt="插入视频文件" width="23" height="22" border="0"></a> <a href="javascript:ra()"><img src="images/ubbcode/bb_ra.gif" alt="插入RA音频文件" border="0"></a> <a href="javascript:rm()"><img src="images/ubbcode/bb_rm.gif" alt="插入RM视频文件" border="0"></a> <a href="javascript:page()"><img src="images/ubbcode/bb_page.gif" alt="插入分页标识符" width="23" height="22" border="0"></a></td>
              </tr>
            </table>
              <table width="98%" border="0" cellpadding="0" cellspacing="0">
                <tr valign="top">
                  <td><textarea name="message" style="width:100%" rows="18" wrap="VIRTUAL" id="Message" onSelect="javascript: storeCaret(this);" onClick="javascript: storeCaret(this);" onKeyDown="javascript: ctlent();" onKeyUp="javascript: storeCaret(this);"></textarea></td>
                </tr>
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td align="right"><b>附件：</b></td>
            <td><iframe border="0" frameBorder="0" frameSpacing="0" height="23" marginHeight="0" marginWidth="0" noResize scrolling="no" width="100%" vspale="0" src="attachment.asp"></iframe></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td align="right"><b>引用：</b></td>
            <td><input name="log_Quote" type="text" id="log_Quote" size="90"><br><span class="smalltxt">要引用另一个用户的网络日志项，请输入以上网络日志项的引用通告 URL。用分号分隔多个地址</span></td>
          </tr>
          <tr align="center" bgcolor="#FFFFFF">
            <td colspan="2"><input name="log_CateID" type="hidden" id="log_CateID" value="<%=Request.Form("log_CateID")%>">
            <input name="PostBlog" type="hidden" value="PostIT"><input name="topicsubmit" type="submit" value=" 提交日志 [可按 Ctrl+Enter 发布] " onClick="this.disabled=true;document.input.submit();">
&nbsp;            <input name="L_Reset" type="reset" id="L_Reset" value=" 重置日志 "></td>
          </tr>
        </form>
    </table></td>
  </tr>
</table><%Else%>
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
    <td width="128" align="center" nowrap><b>发表新日志<br>
          <br>
    <font color="#FF0000">保存日志</font></b></td>
    <td align="center" valign="top"><div class="msg_head">保存日志</div>
            <div class="msg_content"><%IF Request.Form("message")=Empty OR Request.Form("log_Title")=Empty Then
				Response.Write("必须填写日志内容<br><a href='javascript:history.go(-1);'>请返回重新填写</a>")
			Else
				dim Log_Title,log_Content,log_From,log_FromURL,log_CateID2,log_Intro,log_Quote,log_DisSM,log_DisUBB,log_DisIMG,log_AutoURL,log_IsShow,log_AutoKEY,log_IsTop,log_DisComment,log_Weather,log_tags
				log_Title=CheckStr(Request.Form("log_Title"))
				log_Content=CheckStr(Request.Form("message"))
				If InStr(log_Content,"[#splitline#]")>0 Then
				log_Intro=HTMLEncode(Left(log_Content,InStr(log_Content,"[#splitline#]")-1))
				Else
				log_Intro=SplitLines(HTMLEncode(log_Content),4)
				End If
				log_From=CheckStr(Request.Form("log_From"))
				log_FromURL=CheckStr(Request.Form("log_FromURL"))
				log_CateID2=Request.Form("log_CateID")
				log_Quote=CheckStr(Request.Form("log_Quote"))
				log_DisSM=Request.Form("log_DisSM")
				log_DisUBB=Request.Form("log_DisUBB")
				log_DisIMG=Request.Form("log_DisIMG")
				log_AutoURL=Request.Form("log_AutoURL")
				log_AutoKEY=Request.Form("log_AutoKEY")
				log_Weather=Request.Form("log_Weather")
				log_tags=Request.Form("log_tags")
				IF Request.Form("log_IsTop")="1" Then
					log_IsTop=1
				Else
					log_IsTop=0
				End IF
				IF Request.Form("log_IsShow")="0" Then
					log_IsShow=1
				Else
					log_IsShow=0
				End IF
				If Request.Form("log_DisComment")="1" Then
					log_DisComment=1
				Else
					log_DisComment=0
				End IF
				IF log_DisSM=Empty Then log_DisSM=0
				IF log_DisUBB=Empty Then log_DisUBB=0
				IF log_DisIMG=Empty Then log_DisIMG=0
				IF log_AutoURL=Empty Then log_AutoURL=0
				IF log_AutoKEY=Empty Then log_AutoKEY=0

				Conn.ExeCute("INSERT INTO blog_Content(log_CateID,log_Title,log_Author,log_Intro,log_Content,log_From,log_FromURL,log_Quote,log_DisSM,log_DisUBB,log_DisIMG,log_AutoURL,log_AutoKEY,log_IsShow,log_IsTop,log_DisComment,log_Weather) VALUES ("&log_CateID2&",'"&log_Title&"','"&memName&"','"&log_Intro&"','"&log_Content&"','"&log_From&"','"&log_FromURL&"','"&log_Quote&"',"&log_DisSM&","&log_DisUBB&","&log_DisIMG&","&log_AutoURL&","&log_AutoKEY&","&log_IsShow&","&log_IsTop&","&log_DisComment&",'"&log_Weather&"')")
				
				Dim PostLogID
				PostLogID=Conn.ExeCute("SELECT TOP 1 log_ID FROM blog_Content ORDER BY log_ID DESC")(0)
				Conn.ExeCute("UPDATE blog_Member SET mem_PostLogs=mem_PostLogs+1 WHERE mem_Name='"&memName&"'")
				Conn.ExeCute("UPDATE blog_Info SET blog_LogNums=blog_LogNums+1")
				if trim(log_tags)<>empty then
				log_tags=Split(log_tags,";")
				Dim tabs,ts
				For tabs=0 to Ubound(log_tags)
				    Sql="Select TagName from blog_tags where TagName='"&log_tags(tabs)&"'"
					Set ts=Server.CreateObject("Adodb.Recordset")
					ts.open sql,conn,1,1
					If ts.eof and ts.bof then
					   conn.execute ("insert into blog_tags(TagName) values('"&log_tags(tabs)&"')")
					Else
					   conn.execute ("update blog_tags set TagBlogCount=TagBlogCount+1 where tagName='"&log_tags(tabs)&"'")
					End If
					conn.execute ("insert into blog_tag(blog_id,tagsName) values("&PostLogID&",'"&log_tags(tabs)&"')")
					ts.close
				    Set ts=Nothing
				Next
				End IF
				SQLQueryNums=SQLQueryNums+4
				Response.Write("<br>填写日志成功<br><br><a href=""default.asp"">点击返回首页</a><br><br><a href='blogview.asp?logID="&PostLogID&"'>或者返回你所发表的日志</a><br><br>或者等待3秒后自动返回你所发表的日志<meta http-equiv='refresh' content='3;url=blogview.asp?logID="&PostLogID&"'><br><br>")
				IF log_Quote<>Empty And log_IsShow=True Then
					Dim log_QuoteEvery,log_QuoteArr
					log_QuoteArr=Split(log_Quote,";")
					For Each log_QuoteEvery In log_QuoteArr
						Dim TBlogID,TBlogTmp,tbCode,tbMessage
						TBlogID=Conn.Execute("SELECT TOP 1 log_ID FROM blog_Content WHERE log_Author='"&memName&"' AND log_Quote='"&log_Quote&"' ORDER BY log_PostTime DESC")(0)
						SQLQueryNums=SQLQueryNums+1
						If TBlogID>0 then
							TBlogTmp=Split(Trackback(Trim(log_QuoteEvery), siteURL&"/blogview.asp?logID="&TBlogID, log_Title, CutStr(log_Intro,252), siteName),"$$")
							tbCode=TBlogTmp(0)
							tbMessage=TBlogTmp(1)
							If tbCode="0" Then 
								tbMessage="<br><font color=""red"">"&tbMessage&"</font>" 
							Else 
								tbMessage="<br><font color=""blue"">"&tbMessage&"</font>"
							End If
						End if
					Next
				End IF
			End IF%></div></td>
  </tr>
</table><%End IF
End IF%>
<!--#include file="footer.asp" -->