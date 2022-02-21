<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
    <td width="160" valign="top" bgcolor="#C8C7BD" nowrap><%
	Call MemberCenter
	Response.Write("<br>")
	Call SiteInfo
	Response.Write("<br>")
	Call blogSearch%><br>
	</td>
    <td width="100%" valign="top" bgcolor="#FFFFFF"><%
	If memName<>Empty Then
		Dim Fav_Action,Fav_Message
		Fav_Action=Trim(Request.QueryString("action"))
		If Fav_Action="DelFav" Then
			If IsInteger(Request.QueryString("favID"))=False Then
				Fav_Message=Fav_Message&"出现错误：无效的参数"
			Else
				Dim Fav_Delete
				SET Fav_Delete=Server.CreateObject("ADODB.Recordset")
				SQL="SELECT * FROM blog_Favorite WHERE fav_ID="&Request.QueryString("favID")
				Fav_Delete.Open SQL,Conn,1,3
				SQLQueryNums=SQLQueryNums+1
				If Fav_Delete.EOF And Fav_Delete.BOF Then
					Fav_Message=Fav_Message&"出现错误：你所要删除的书签不存在"
				Else
					IF Not (memStatus="SupAdmin" OR memName=Fav_Delete("fav_Author")) Then
						Fav_Message=Fav_Message&"出现错误：你没有权限删除此书签"
					Else
						Fav_Delete.Delete
						Fav_Delete.Update
						Fav_Message=Fav_Message&"删除成功：你所选择的书签已经删除"
					End If
				End If
				Fav_Delete.Close
				Set Fav_Delete=Nothing
			End If
			Response.Write("<div class=""msg_head""></div><div class=""msg_content"">"&Fav_Message&"<br><a href=""favorite.asp"">点击返回</a></div>")
		ElseIf Fav_Action="AddFav" Then
			If Request.Form("favsite_Name")=Empty Or Request.Form("favsite_URL")=Empty Then
				Fav_Message=Fav_Message&"出现错误：请将信息填写完整"
			Else
				Fav_Message=Fav_Message&"添加成功：你所添加的书签添加成功"
				Conn.ExeCute("INSERT INTO blog_Favorite(fav_Name,fav_URL,fav_Info,fav_Author) VALUES ('"&CheckStr(Request.Form("favsite_Name"))&"','"&Request.Form("favsite_URL")&"','"&CheckStr(Request.Form("favsite_Info"))&"','"&memName&"')")
				SQLQueryNums=SQLQueryNums+1
			End If
			Response.Write("<div class=""msg_head""></div><div class=""msg_content"">"&Fav_Message&"<br><a href=""favorite.asp"">点击返回</a></div>")
		Else%><div class="content_head"><strong><%=memName%> 的个人书签</strong></div><div class="content_main"><%Dim Fav_List
	Set Fav_List=Server.CreateObject("ADODB.RecordSet")
	SQL="SELECT * FROM blog_Favorite WHERE fav_Author='"&memName&"' ORDER BY fav_PostTime ASC"
	Fav_List.Open SQL,Conn,1,1
	SQLQueryNums=SQLQueryNums+1
	If Fav_List.EOF And Fav_List.BOF Then
		Response.Write("<div class=""message"">暂时没有个人书签</div>")
	Else
		Do While Not Fav_List.EOF
			Response.Write("<span class=""hyperlink""><a href=""favorite.asp?action=DelFav&FavID="&Fav_List("Fav_ID")&""" title=""删除此个人书签"" onClick=""winconfirm('你真的要删除这个书签吗？','favorite.asp?action=DelFav&FavID="&Fav_List("Fav_ID")&"'); return false""><b><font color=""#FF0000"">×</font></b></a>&nbsp;|&nbsp;<a href="""&Fav_List("Fav_URL")&""" target=""_blank"" alt="""&Fav_List("Fav_Info")&""">"&Fav_List("Fav_Name")&"</a></span>")
			Fav_List.MoveNext
		Loop
	End If
	Fav_List.Close
	Set Fav_List=Nothing
	%></div>
	<div class="content_head"><strong>添加个人书签</strong> - 为了大家，请不要提交非法信息</div><div class="content_main"><table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	  <form name="blogFavAdd" method="post" action="favorite.asp?action=AddFav">
      <tr bgcolor="#FFFFFF">
        <td width="12%" align="center" nowrap>站点名称：</td>
        <td width="54%"><input name="favsite_Name" type="text" id="favsite_Name" size="32"> 
        <strong>必填</strong></td>
        <td width="34%">站点介绍：</td>
        </tr>
      <tr bgcolor="#EFEFEF">
        <td align="center" nowrap bgcolor="#FFFFFF">站点地址：</td>
        <td bgcolor="#FFFFFF"><input name="favsite_URL" type="text" id="favsite_URL" value="http://" size="32">
          <strong>必填</strong></td>
        <td rowspan="2" bgcolor="#FFFFFF"><textarea name="favsite_Info" cols="40" rows="3" wrap="VIRTUAL" id="favsite_Info"></textarea></td>
      </tr>
      <tr bgcolor="#EFEFEF">
        <td align="center" nowrap bgcolor="#FFFFFF"></td>
        <td bgcolor="#FFFFFF"><input type="submit" name="Submit" value=" 提交书签 "></td>
        </tr></form>
    </table>
	</div><%
	End If
	Else
		Response.Write("<div class=""message""><strong>未注册用户不能使用网络书签功能</strong></div>")
	End If%></td>
 </tr>
</table>
<!--#include file="footer.asp" -->