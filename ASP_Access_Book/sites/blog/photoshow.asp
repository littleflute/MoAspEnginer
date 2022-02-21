<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/ubbcode.asp" -->
<!--#include file="include/md5code.asp" -->
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
  <tr>
	<td>
<%IF Request.QueryString("action")="postcomm" Then
	Dim ph_commID,msg_Title,msg_Content
	ph_commID=Request.Form("ph_commID")
	dim server_v1,server_v2
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
	   Response.write"<script>alert(""你提交的路径有误，禁止从站点外部提交数据请不要乱该参数！"");history.go(-1);</script>"
	   Response.end
	elseIf IsInteger(ph_commID)=False Then
		Response.Write("<script>alert(""参数出现错误"");history.go(-1);</script>")
		Response.End
	ElseIf (memStatus<>"SupAdmin" And memStatus<>"Admin") And DateDiff("s",Request.Cookies(CookieName)("memLastPost"),Now())<10 Then
		Response.Write("<script>alert(""你发表评论速度太快了,请10秒后再操作!"");history.go(-1);</script>")
		Response.End
	Else
	  	Dim comm_LogQuery,comm_LogISOK
		Set comm_LogQuery=Conn.ExeCute("SELECT ph_DisComm FROM Photo WHERE ph_ID="&ph_commID&"")
		IF comm_LogQuery.EOF AND comm_LogQuery.BOF Then
			comm_LogISOK=1
		Else
			IF comm_LogQuery(0)=True Then
				comm_LogISOK=2
			End IF
		End IF
		Set comm_LogQuery=Nothing
		Dim comm_AllreadyMem,comm_AllreadyMemErr
		Set comm_AllreadyMem=Server.CreateObject("ADODB.RecordSet")
		SQL="SELECT mem_Name,mem_Password,mem_Status,mem_LastIP FROM blog_Member WHERE mem_Name='"&CheckStr(Request.Form("username"))&"'"
		comm_AllreadyMem.Open SQL,Conn,1,3
		SQLQueryNums=SQLQueryNums+1
		IF comm_AllreadyMem.EOF AND comm_AllreadyMem.BOF Then
			comm_AllreadyMemErr=0
		ElseIF comm_AllreadyMem("mem_Password")=MD5(CheckStr(Request.Form("mem_Password"))) Then
			Response.Cookies(CookieName)("memName")=comm_AllreadyMem("mem_Name")
			Response.Cookies(CookieName)("memPassword")=comm_AllreadyMem("mem_Password")
			Response.Cookies(CookieName)("memStatus")=comm_AllreadyMem("mem_Status")
			memName=comm_AllreadyMem("mem_Name")
			comm_AllreadyMem("mem_LastIP")=Guest_IP
			comm_AllreadyMem.Update
			comm_AllreadyMemErr=2
		Else
			comm_AllreadyMemErr=1
		End IF
		comm_AllreadyMem.Close
		Set comm_AllreadyMem=Nothing
		IF CheckStr(Request.Form("message"))=Empty OR CheckStr(Request.Form("username"))=Empty Then
			Response.Write("<script>alert(""请将必须信息填写完整!"");history.go(-1);</script>")
			Response.End
		ElseIF Len(CheckStr(Request.Form("message")))>1008 OR Len(CheckStr(Request.Form("message")))<5 Then
			Response.Write("<script>alert(""评论内容不能少于5个字符或超过1008个字符!"");history.go(-1);</script>")
			Response.End
		ElseIF Len(CheckStr(Request.Form("username")))>12 OR Len(CheckStr(Request.Form("username")))<2 Then
			Response.Write("<script>alert(""用户名长度不能少于2或者大于10个字符，5个汉字!"");history.go(-1);</script>")
			Response.End
		ElseIF IsValidUserName(CheckStr(Request.Form("username")))=False Then
			Response.Write("<script>alert(""用户名中含有非法字符!"");history.go(-1);</script>")
			Response.End
		ElseIF memName=Empty AND comm_AllreadyMemErr=1 Then
			Response.Write("<script>alert(""很遗憾，你所使用的用户名已经注册!"");history.go(-1);</script>")
			Response.End
		ElseIF comm_LogISOK=1 Then
			Response.Write("<script>alert(""对不起，你所要评论的不存在或已删除!"");history.go(-1);</script>")
			Response.End
		ElseIF Not(memStatus="SupAdmin" OR memStatus="Admin") AND comm_LogISOK=2 Then
			Response.Write("<script>alert(""对不起，你所评论的日志不允许发表评论!"");history.go(-1);</script>")
			Response.End
		'屏蔽无聊链接(2002/03/17)
		ElseIF Strurls(Request.Form("message"),"[url")>MaxUrl or Strurls(Request.Form("message"),"http://")>MaxHttp then 
			Response.Write("<script>alert(""您输入的内容有非法链接\n只能有两个连接\n请返回重新输入"");history.go(-1);</script>")
			Response.End
		Else
			Dim comm_Content,comm_memName,comm_DisSM,comm_DisUBB,comm_DisIMG,comm_AutoURL,comm_AutoKEY,comm_memFace,comm_Hide
			comm_Content=CheckStr(Request.Form("message"))
			comm_memName=CheckStr(Request.Form("username"))
			comm_Hide=Request.Form("hidden_message")
			'comm_memFace=CheckStr(Request.Form("userface"))
			IF comm_Hide=Empty Then comm_Hide=1
			IF memName=Empty And comm_AllreadyMemErr<>2 Then
				Dim SaveMem,comm_memPassword
				SaveMem=Request.Form("SaveMem")
				comm_memPassword=MD5(CheckStr(Request.Form("mem_Password")))
				IF SaveMem=1 Then
					Conn.ExeCute("INSERT INTO blog_Member(mem_Name,mem_Password,mem_LastIP) VALUES ('"&comm_memName&"','"&comm_memPassword&"','"&Guest_IP&"')")
					Conn.ExeCute("UPDATE blog_Info SET blog_MemNums=blog_MemNums+1")
					SQLQueryNums=SQLQueryNums+2
					Response.Cookies(CookieName)("memName")=comm_memName
					Response.Cookies(CookieName)("memPassword")=comm_memPassword
					Response.Cookies(CookieName)("memStatus")="Member"
				End IF
				Conn.ExeCute("INSERT INTO photo_Comment(ph_ID,comm_Content,comm_Author,comm_Hide,comm_PostIP) VALUES ("&ph_commID&",'"&comm_Content&"','"&comm_memName&"',"&comm_Hide&",'"&Guest_IP&"')")
				SQLQueryNums=SQLQueryNums+1
			Else
				Conn.ExeCute("INSERT INTO photo_Comment(ph_ID,comm_Content,comm_Author,comm_Hide,comm_PostIP) VALUES ("&ph_commID&",'"&comm_Content&"','"&memName&"',"&comm_Hide&",'"&Guest_IP&"')")
				SQLQueryNums=SQLQueryNums+1
			End IF
			Application.Lock
			'Application.Contents(CookieName&"_Photo_LastComm") = ""
			Application.UnLock
			Conn.ExeCute("UPDATE Photo SET ph_Comments=ph_Comments+1 WHERE ph_ID="&ph_commID&"")
			SQLQueryNums=SQLQueryNums+1
			Response.Cookies(CookieName)("memLastpost")=Now()
			Response.Write("<script>alert(""发表成功,谢谢参与!"");document.location.href=""photoshow.asp?photoID="&ph_commID&"#comment"";</script>")
		End If
	End If
ElseIF Request.QueryString("action")="delecomm" Then
	IF IsInteger(Request.QueryString("commID"))=False OR IsInteger(Request.QueryString("photoID"))=False Then
		Response.Write("<script>alert(""参数出现错误"");history.go(-1);</script>")
		Response.End
	Else
		Dim ph_AuthorQuery
		Set ph_AuthorQuery=Conn.ExeCute("SELECT ph_Author FROM Photo WHERE ph_ID="&CheckStr(Request.QueryString("photoID")))
		SQLQueryNums=SQLQueryNums+1
		IF ph_AuthorQuery.EOF AND ph_AuthorQuery.BOF Then
			Response.Write("<script>alert(""参数出现错误"");history.go(-1);</script>")
			Response.End
		Else
			IF Not (memStatus="SupAdmin" OR (memStatus="Admin" And memName=ph_AuthorQuery(0))) Then
				Response.Write("<script>alert(""你没有权限删除!"");history.go(-1);</script>")
			Else
				Dim dele_Comm
				Set dele_Comm=Conn.ExeCute("SELECT ph_ID,comm_Author FROM photo_Comment WHERE comm_ID="&CheckStr(Request.QueryString("commID")))
				SQLQueryNums=SQLQueryNums+1
				IF dele_Comm.EOF AND dele_Comm.BOF Then
					Response.Write("<script>alert(""没有找到指定数据!"");history.go(-1);</script>")
				Else
					Conn.ExeCute("UPDATE photo SET ph_Comments=ph_Comments-1 WHERE ph_ID="&dele_Comm("ph_ID"))
					Conn.Execute("DELETE  FROM photo_Comment WHERE comm_ID="&CheckStr(Request.QueryString("commID")))
					SQLQueryNums=SQLQueryNums+4
					Application.Lock
					'Application.Contents(CookieName&"_blog_DelPhComm") = ""
					Application.UnLock
					Response.Write("<script>alert(""删除成功!"");document.location.href=""photoshow.asp?photoID="&CheckStr(Request.QueryString("photoID"))&""";</script>")
				End IF
				Set dele_Comm=Nothing
			End IF
		End IF
		Set ph_AuthorQuery=Nothing
	End IF
ElseIf Request.QueryString("action")="postvote" Then
	Dim voteNums,formV,VoteNum0,VoteNum1,VoteNum2,VoteNum3,VoteNum4
	ph_commID=Request.Form("ph_commID")
	voteNums=Request.Form("voteNums")
	voteNum0=Request.Form("voteNum0")
	voteNum1=Request.Form("voteNum1")
	voteNum2=Request.Form("voteNum2")
	voteNum3=Request.Form("voteNum3")
	voteNum4=Request.Form("voteNum4")
	Dim ph_Vote
	select case voteNums
		case "1"
			voteNum0 = (cint(voteNum0)+1)
		case "2"
			voteNum1 = (cint(voteNum1)+1)
		case "3"
			voteNum2 = (cint(voteNum2)+1)
		case "4"
			voteNum3 = (cint(voteNum3)+1)
		case "5"
			voteNum4 = (cint(voteNum4)+1)
	end select
	ph_Vote = voteNum0 & "|" & voteNum1 & "|" & voteNum2 & "|" & voteNum3 & "|" & voteNum4
	If voteNums=Empty Or voteNums=0 Then
		Response.Write("<script>alert(""请点击分数选项!"");history.go(-1);</script>")
		Response.End
	ElseIf DateDiff("s",Request.Cookies(CookieName)("ph_Votepost"&ph_commID&""&ph_commID&""),Now())<43200 Then
		Response.Write("<script>alert(""你已打过分了,谢谢关注!"");history.go(-1);</script>")
		Response.End
	Else
		Conn.ExeCute("UPDATE Photo SET ph_Vote='"&ph_Vote&"' WHERE ph_ID="&ph_commID&"")
		SQLQueryNums=SQLQueryNums+1
		Response.Cookies(CookieName)("ph_Votepost"&ph_commID&""&ph_commID&"")=Now()
		Response.Write("<script>alert(""打分成功,谢谢关注!"");document.location.href=""photoshow.asp?photoID="&ph_commID&"#vote"";</script>")
	End If
Else

	Dim PhotoID
	PhotoID=CheckStr(Trim(Request.QueryString("PhotoID")))
	If Not IsInteger(PhotoID) Then PhotoID=0
	IF PhotoID=Empty Then
		Response.Write("<div align='center'><h4>参数错误,请不要乱提交数据!</h4><br /><br /><a href=""javascript:history.go(-1)"">返回上一页</a> 或 <a href=""default.asp"">返回首页</a></div>")
	Else
		Dim Ph_Show
		Set Ph_Show=Server.CreateObject("ADODB.RecordSet")
		SQL="SELECT * FROM Photo WHERE ph_ID="&PhotoID&""
		Ph_Show.Open SQL,Conn,1,3
		If Ph_Show.BOF AND Ph_Show.EOF Then
			Response.Write("<div align='center'><h4>参数错误,没有找到数据!</h4><br /><br /><a href=""javascript:history.go(-1)"">返回上一页</a> 或 <a href=""default.asp"">返回首页</a></div>")
		Else
			Dim phid,phCateIDz,phvote,Ph_Images,n,TotalNum,photoImage,photo_Prev,photo_Next
			phid=Ph_Show("ph_ID")
			phCateID=Ph_Show("ph_CateID")
			phvote=split(Ph_Show("ph_Vote"),"|")
			Ph_Images=split(Ph_Show("ph_Image"),vbcrlf)
			TotalNum = 0
			'计算投票总数
			for n = 0 to 4
				TotalNum = phvote(n) + TotalNum
			next
			Dim ph_Next
			Set ph_Next=Conn.Execute("SELECT TOP 1 ph_ID,ph_Name FROM Photo WHERE ph_ID>"&PhotoID&" ORDER BY ph_ID ASC")
			SQLQueryNums=SQLQueryNums+1
			If ph_Next.EOF AND ph_Next.BOF Then
				photo_Next=""
			Else
				photo_Next="&nbsp;&nbsp;&nbsp;&nbsp;<a href=""photoshow.asp?photoID="&ph_Next(0)&""" title=""下一张"">"&ph_Next(1)&"</a> >>"
			End If
			ph_Next.close
			Set ph_Next=Nothing
			
			Dim ph_Prev
			Set ph_Prev=Conn.Execute("SELECT TOP 1 ph_ID,ph_Name FROM Photo WHERE ph_ID<"&PhotoID&" ORDER BY ph_ID DESC")
			SQLQueryNums=SQLQueryNums+1
			If ph_Prev.EOF AND ph_Prev.BOF Then
				photo_Prev=""
			Else
				photo_Prev="<< <a href=""photoshow.asp?photoID="&ph_Prev(0)&""" title=""上一张"">"&ph_Prev(1)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;"
			End If
			ph_Prev.close
			Set ph_Prev=Nothing

			dim ph_cate,ph_CateName
			Set ph_cate=Conn.Execute("SELECT cate_Name FROM photo_Cate WHERE cate_ID="&phCateID&"")
			IF ph_cate.EOF AND ph_cate.BOF Then
				ph_CateName="&nbsp;"
			Else
				ph_CateName="&nbsp;&nbsp;&nbsp;&nbsp;所在分类：<a href='photo.asp?cateID="&phCateID&"'><strong>"&ph_cate(0)&"</strong></a>"
			End IF
			ph_cate.close
			Set ph_cate=Nothing%>
			<div id="ph_main">
			<div id="ph_main_bg">
				<!-- 循环组图 -->
				<%Dim Ph_Img
					For Each Ph_Img In Ph_Images
				If Not (Right(Ph_Img, 1) = "@") Then Ph_Img = Ph_Img & "@"
						Dim ph_Itexts
						ph_Itexts=split(""&Ph_Img&"","@")
						Dim Ph_Image,Ph_Text
						Ph_Image=ph_Itexts(0)
						Ph_Text=ph_Itexts(1)
				%>
					<table border="0" align="center" cellspacing="0" cellpadding="0" style="margin-top:10px; margin-bottom:10px;">
					  <tr>
						<td id="photo_show">
							<!-- 加载图片防止宽度大于680 -->
							<img onload='javascript:if(this.width>680)this.width=680;' src="<%=""&Ph_Image&""%>?v=<%=""&phid&""%>" border="0" alt="" />
						</td>
					  </tr>
					  <tr>
						<td style="text-indent: 20px;"><%=""&Ph_Text&""%></td>
					  </tr>
					</table>
				    <%Next%>
				<!-- 循环组图结束 -->
			    <table width="98%" style="margin-bottom: 10px;" border="0" align="center" cellpadding="6" cellspacing="6" bgcolor="#FFFFFF">
                  <tr bgcolor="#C8C7BD">
                    <td width="34%" align="left" valign="top">
						<p>图片名称：<strong><%=""&Ph_Show("ph_Name")&""%><br>
					  <br>
						</strong>浏览/评论次数：<%=""&Ph_Show("ph_Views")&""%>/<%=""&Ph_Show("ph_Comments")&""%><br>
					  发布作者：<a href="member.asp?action=view&memName=<%=""&Server.URLEncode(Ph_Show("ph_Author"))&""%>"><%=""&Ph_Show("ph_Author")&""%><br>
					  </a>添加时间：<%=""&Ph_Show("ph_PostTime")&""%>
							<%IF (memStatus="Admin" AND memName=CommAuthor) OR memStatus="SupAdmin" Then
								Response.Write("&nbsp;<a href=""photoedit.asp?photoID="&phid&""" title=""编辑"" target=""_blank""><img src='images/icon_edit.gif' border='0' align='absmiddle'></a>")
							End IF%>
					  </p>
					</td>
					<td width="66%" valign="top">
						<p>图片简介：<br>
						  <%Response.Write(UbbCode(HTMLEncode(Ph_Show("ph_Remark")),0,0,1,1,0))%>
						</p>
					</td>
                  </tr>
                  <tr>
                    <td colspan="2" bgcolor="#C8C7BD"><%=""&photo_Prev&""%> <%=""&photo_Next&""%> <%=""&ph_CateName&""%></td>
                  </tr>
              </table>
		      <br />
		</div>
	  </div>
			<div id="ph_main">
		<div id="ph_main_bg">
		  <div id="photo_left">
				<%if Ph_Show("Ph_DisVote")=false then%>
				<a name="vote"></a>
				<br>
				<%end if%>
			  <!-- 评论开始 -->
		  </div>
			<div id="photo_right">
			<%if Ph_Show("Ph_DisComm")=false then%>
				<%Dim CurPage,Url_Add
				Url_Add="?photoID="&PhotoID&"&"
				If CheckStr(Request.QueryString("Page"))<>Empty Then
					Curpage=CheckStr(Request.QueryString("Page"))
					If IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
					Else
						Curpage=1
					End If
					Dim ph_Comment
					Set ph_Comment=Server.CreateObject("Adodb.RecordSet")
					SQL="SELECT * FROM photo_Comment WHERE ph_ID="&PhotoID&" ORDER BY comm_PostTime DESC"
					ph_Comment.Open SQL,Conn,1,1
					SQLQueryNums=SQLQueryNums+1
					IF ph_Comment.EOF AND ph_Comment.BOF Then%>
						<table width="100%" border="0" cellpadding="6" cellspacing="1" class="photo_bg">
						 <tr>
							<td><strong>所有评论</strong></td>
						 </tr>
						 <tr bgcolor="#ffffff">
							<td>暂无</td>
						 </tr>
						</table>
					<%Else
						Dim Comm_Nums,MultiPages,PageCount
						ph_Comment.PageSize=5
						ph_Comment.AbsolutePage=CurPage
						Comm_Nums=ph_Comment.RecordCount
						MultiPages=""&MultiPage(Comm_Nums,5,CurPage,Url_Add)&""%>
						<a name="comment"></a>
						<table width="100%" border="0" align="center" cellpadding="6" cellspacing="1" class="photo_bg">
						  <tr>
							<td width="20%"><strong>所有评论</strong></td>
							<td align="right"><%=""&MultiPages&""%></td>
						  </tr>
							<%Do Until ph_Comment.EOF OR PageCount=5
							Dim CommID,CommAuthor,CommContent,CommHide,commPostIP
							CommID=ph_Comment("comm_ID")
							CommAuthor=ph_Comment("comm_Author")
							CommContent=ph_Comment("comm_Content")
							CommHide=ph_Comment("comm_Hide")
							commPostIP=ph_Comment("comm_PostIP")%>
						  <tr bgcolor="#ffffff">
							<td colspan="2"><%If CommHide= 1 OR (memStatus="SupAdmin" OR memStatus="Admin" OR memName=CommAuthor) Then
								Response.Write(UbbCode(HTMLEncode(CommContent),0,0,0,1,0))
							Else
								Response.Write("<font color=red>这是一篇隐藏评论,只有管理员才能查看!</font>")
							End If%>
							  <br>
							<strong><a href="member.asp?action=view&memName=<%=""&Server.URLEncode(CommAuthor)&""%>" target="new"><%=""&CommAuthor&""%></a></strong>&nbsp;&nbsp;
								发表于&nbsp;&nbsp;<%=""&DateToStr(ph_Comment("comm_PostTime"),"Y-m-d H:I A")&""%>								<%IF (memStatus="Admin" AND memName=CommAuthor) OR memStatus="SupAdmin" Then
									Response.Write("&nbsp;<a href=""photoedit.asp?action=editcomm&commID="&CommID&""" title=""编辑评论"" target=""_blank""><img src='images/icon_edit.gif' border='0' align='absmiddle'></a>")
									Response.Write("&nbsp;<a href=""photoshow.asp?action=delecomm&photoID="&PhotoID&"&commID="&CommID&""" title=""删除评论"" onClick=""winconfirm('你真的要删除这个评论吗？','photoshow.asp?action=delecomm&photoID="&PhotoID&"&commID="&CommID&"'); return false""><img src=""images/icon_del.gif"" border=""0"" align=""absmiddle"" alt="""" /></a>")
									Response.Write("&nbsp;<a href=""http://whois.pconline.com.cn/whois/?ip="&commPostIP&""" title=""点击查看IP来源"" target=""_blank""><img src=""images/ip.gif"" alt="""" border=""0"" align=""absmiddle"" /></a>")
								End IF%>
							</td>
						  </tr>
						<%PageCount=PageCount+1
						ph_Comment.MoveNext
						Loop%>
						</table>
					<%End IF
				ph_Comment.Close
				Set ph_Comment=Nothing
			End if%>
				<%if Ph_Show("Ph_DisComm")=false then%>
				<script language="JavaScript" src="include/ubbcode.js" type="text/javascript"></script>
				<table width="100%" border="0" cellspacing="1" cellpadding="6" class="photo_bg">
				  <tr>
					<td><strong>看图评论</strong>(不用注册也可发表评论，但要输入呢称)</td>
				  </tr>
				  <tr bgcolor="#ffffff">
					<td>
						<form action="photoshow.asp?action=postcomm" method="post" name="inputform" id="inputform">
						<p>
							呢称：
							<%IF memName<>Empty Then%>
								<input name="username" type="text" id="username" value="<%=""&memName&""%>" class="input_bg" size="10" readonly />
							<%Else%>
								<input name="username" type="text" id="username" class="input_bg" size="10" maxlength="12" onMouseOver="this.focus();" />&nbsp;&nbsp;
								密码：
								<input name="mem_Password" type="password" id="mem_Password" class="input_bg" size="12" />&nbsp;&nbsp;
								<input name="SaveMem" type="checkbox" id="SaveMem" value="1" />注册成员?
							<%End IF%>
						</p>
						<p><textarea name="message" style="width: 98%;" rows="6" wrap="VIRTUAL" id="Message" onSelect="javascript: storeCaret(this);" onClick="javascript: storeCaret(this);" onKeyUp="javascript: storeCaret(this);" onKeyDown="javascript: ctlent();"></textarea></p>
						<div style="float: right">不支持html代码&nbsp;&nbsp;&nbsp;&nbsp; 缩放输入框: <span title="放大" style="cursor:hand" onClick="document.inputform.message.rows+=5"><img src="images/icon_ar2.gif" align="absmiddle" border="0" /></span> <span title="缩小" style="cursor:hand" onclick='if(document.inputform.message.rows>=5)document.inputform.message.rows-=5;else return false'><img src="images/icon_al2.gif" align="absmiddle" border="0" /></span></div>
						<p align="left">
						  <input name="hidden_message" type="checkbox" id="hidden_message" value="0" /> 隐藏评论<br>
						  <input name="ph_commID" value="<%=""&PhotoID&""%>" type="hidden" />
							<input type="submit" name="replysubmit" value=" 发表 " onClick="this.disabled=true;document.inputform.submit();" />
							<input name="reset" type="reset" value=" 重置 " />
						</p>
						</form>
					</td>
				  </tr>
				</table>
			<%end if%>
			</div>
		</div>
		</div>
			<%Conn.ExeCute("UPDATE Photo SET ph_Views=ph_Views+1 WHERE ph_ID="&PhotoID&"")
			Ph_Show.UPDATE
			Ph_Show.Close
		End If
	End If
End If%>
    <td>
  <tr>
</table>
<!--#include file="footer.asp" -->