<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/ubbcode.asp" -->
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif" class="wordbreak">
  <tr>
   <td width="160" valign="top" bgcolor="#C8C7BD" nowrap><%Call MemberCenter%></td>
    <td align="center" class="mainbox"> 
      <h3>编辑下载</h3>
      <%IF Request.QueryString("action")<>"postedit" Then
Dim downls_ID
downls_ID=CheckStr(Request.QueryString("downlID"))
If Not IsInteger(downls_ID) Then downls_ID=0
Dim downl_Edit
Set downl_Edit=Conn.Execute("SELECT * FROM blog_Download WHERE downl_ID="&downls_ID&"")
IF downl_Edit.EOF AND downl_Edit.BOF Then%>
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="messagebox">
        <tr>
				
          <td align="center"><b><font color="#FF0000">下载内容不存在</font> </b><br><br><a href="download.asp">点击返回</a></td>
			</tr>
	</table>
<%
Else
	IF Not(("downl_Author"=memName AND memStatus="Admin") OR memStatus="SupAdmin") Then%>
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="messagebox">
        <tr>
				
          <td align="center"><b><font color="#FF0000">你没有权限</font> </b><br><br>
				<a href="default.asp">点击返回主页面</a><br>
            <br>
            <a href="logging.asp">或者重新登陆</a></td>
				</tr>
		</table>
<%Else%>
<script language="JavaScript" src="include/ubbcode.js"></script>
      <form name="input" method="post" action="downledit.asp?action=postedit">
        <table width="100% border="0" align="center" cellpadding="4" cellspacing="1" class="formbox">
          <tr> 
            <td colspan="5" class="formbox-title">编辑下载</td>
          </tr>
          <tr class="formbox-content"> 
            <td align="right" >分类:</td>
            <td colspan="3"><select name="downl_Cate">
                <option value="1" <%If downl_Edit("downl_Cate")="1" then Response.write " selected"%>>文档</option>
                <option value="2"<%If downl_Edit("downl_Cate")="2" then Response.write " selected"%>>墙纸</option>
                <option value="3"<%If downl_Edit("downl_Cate")="3" then Response.write " selected"%>>其他</option>
              </select>&nbsp;&nbsp;文件名: <input name="downl_Name" type="text" size='33' id="downl_Name3" value="<%=downl_Edit("downl_Name")%>">&nbsp;&nbsp;大小:<input name="downl_size" size='5' type="text" id="downl_size" value="<%=downl_Edit("downl_size")%>"></td>
          </tr>
          <tr> 
            <td align="right" valign="middle">来自:</td>
            <td valign="top"><input name="downl_From" type="text" id="downl_From" value="<%=downl_Edit("downl_From")%>"></td>
            <td align="right">来自链接: </td>
            <td valign="top"><input name="downl_FromURL" type="text" id="downl_FromURL" value="<%=downl_Edit("downl_FromURL")%>"></td>
          </tr>
          <tr class="formbox-content"> 
            <td align="right">作者:</td>
            <td><input name="downl_Author" type="text" id="downl_Author5" value="<%=downl_Edit("downl_Author")%>"></td>
            <td align="right">图片链接:</td>
            <td><input name="downl_ImgPath" type="text" id="downl_ImgPath" value="<%=downl_Edit("downl_ImgPath")%>"></td>
          </tr>
          <tr class="formbox-content"> 
            <td align="right">地址1说明:</td>
            <td><input name="downl_Dcomm1" type="text" id="downl_Dcomm1" value="<%=downl_Edit("downl_Dcomm1")%>"></td>
            <td align="right">地址1 URL:</td>
            <td><input name="downl_Dcomm1URL" type="text" id="downl_Dcomm1URL" value="<%=downl_Edit("downl_Dcomm1URL")%>"></td>
          </tr>
          <tr class="formbox-content"> 
            <td align="right">地址2说明:</td>
            <td><input name="downl_Dcomm2" type="text" id="downl_Dcomm2" value="<%=downl_Edit("downl_Dcomm2")%>"></td>
            <td align="right">地址2 URL:</td>
            <td><input name="downl_Dcomm2URL" type="text" id="downl_Dcomm2URL" value="<%=downl_Edit("downl_Dcomm2URL")%>"></td>
          </tr>
          <tr class="formbox-content"> 
            <td align="right">地址3说明:</td>
            <td><input name="downl_Dcomm3" type="text" id="downl_Dcomm3" value="<%=downl_Edit("downl_Dcomm3")%>"></td>
            <td align="right">地址3 URL:</td>
            <td> <input name="downl_Dcomm3URL" type="text" id="downl_Dcomm3URL" value="<%=downl_Edit("downl_Dcomm3URL")%>"></td>
          </tr>
          <tr class="formbox-content"> 
            <td align="right">地址4说明:</td>
            <td><input name="downl_Dcomm4" type="text" id="downl_Dcomm4" value="<%=downl_Edit("downl_Dcomm4")%>"></td>
            <td align="right">地址4 URL:</td>
            <td><input name="downl_Dcomm4URL" type="text" id="downl_Dcomm4URL" value="<%=downl_Edit("downl_Dcomm4URL")%>"></td>
          </tr>
          <tr class='formbox-content'> 
            <td align="right">文件或图片地址:</td>
            <td colspan="4"><input type='text' size='55' id='downl_site' name='downl_site'>
            </td>
          </tr>
          <%if CheckRights(memStatus,"Upload")>0 then%>
          <tr> 
            <td align="right" class="formbox-content">附件:</td>
            <td colspan="3" class="formbox-content"><iframe border="0" frameBorder="0" frameSpacing="0" height="21" marginHeight="0" marginWidth="0" noResize scrolling="no" width="100%" vspale="0" src="attachment.asp"></iframe></td>
          </tr>
          <%end if%>
          <tr class='formbox-content'> 
            <td align="right" valign="top">说明: </td>
            <td colspan="3" align="left" valign="top"><textarea name="downl_Comment" cols="70" rows="6" wrap="VIRTUAL" id="downl_Comment"><%=downl_Edit("downl_Comment")%></textarea></td>
          </tr>
          <tr align="center"> 
            <td colspan="4"><input name="downl_ID" type="hidden" id=downl_ID" value="<%=downl_Edit("downl_ID")%>"> 
              <input name="editsubmit" type="submit" value=" 确定编辑 " onClick="this.disabled=true;document.input.submit();"> 
              &nbsp; <input name="L_Reset" type="button" id="L_Reset6" value=" 重置 " onClick="javascript:history.go(-1);"></td>
          </tr>
        </table>
      </form>
      <%
End IF
End IF
Set downl_Edit=Nothing
Else%>
      <br>
      <table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="messagebox">
        <tr>
          <td align="center">保存</td>
        </tr>
          <tr>
            
          <td align="center" valign="middle" class="messagebox-content"> 
            <%
			IF Trim(Request.Form("downl_Name"))=Empty OR Trim(Request.Form("downl_From"))=Empty OR Trim(Request.Form("downl_Dcomm1"))=Empty OR Trim(Request.Form("downl_Dcomm1URL"))=Empty THEN
				Response.Write("有些必填项没有被填充。<br><br><a href='javascript:history.go(-1);'>返回</a>")
			Else
				Dim downl_ID,downl_Cate,downl_Name,downl_Author,downl_From,downl_size,downl_FromURL,downl_ImgPath,downl_Comment,downl_Dcomm1,downl_Dcomm1URL,downl_Dcomm2,downl_Dcomm2URL,downl_Dcomm3,downl_Dcomm3URL,downl_Dcomm4,downl_Dcomm4URL
		        downl_ID=Request.Form("downl_ID")
		        downl_Cate=Request.Form("downl_Cate")
		        downl_Name=CheckStr(Request.Form("downl_Name"))
		        downl_Author=Request.Form("downl_Author")
		        downl_From=Request.Form("downl_From")
		        downl_size=Request.Form("downl_size")
		        downl_FromURL=Request.Form("downl_FromURL")
		        downl_ImgPath=Request.Form("downl_ImgPath")
		        downl_Comment=Request.Form("downl_Comment")
		        downl_Dcomm1=Request.Form("downl_Dcomm1")
		        downl_Dcomm1URL=Request.Form("downl_Dcomm1URL")
		        downl_Dcomm2=Request.Form("downl_Dcomm2")
		        downl_Dcomm2URL=Request.Form("downl_Dcomm2URL")
		        downl_Dcomm3=Request.Form("downl_Dcomm3")
		        downl_Dcomm3URL=Request.Form("downl_Dcomm3URL")
		        downl_Dcomm4=Request.Form("downl_Dcomm4")
		        downl_Dcomm4URL=Request.Form("downl_Dcomm4URL")
				Dim downl_MoveToSQL
				Conn.ExeCute("UPDATE blog_Download Set downl_Cate='"&downl_Cate&"',downl_Name='"&downl_Name&"',downl_Author='"&downl_Author&"',downl_From='"&downl_From&"',downl_size='"&downl_size&"',downl_FromURL='"&downl_FromURL&"',downl_ImgPath='"&downl_ImgPath&"',downl_Comment='"&downl_Comment&"',downl_Dcomm1='"&downl_Dcomm1&"',downl_Dcomm1URL='"&downl_Dcomm1URL&"',downl_Dcomm2='"&downl_Dcomm2&"',downl_Dcomm2URL='"&downl_Dcomm2URL&"',downl_Dcomm3='"&downl_Dcomm3&"',downl_Dcomm3URL='"&downl_Dcomm3URL&"',downl_Dcomm4='"&downl_Dcomm4&"',downl_Dcomm4URL='"&downl_Dcomm4URL&"' WHERE downl_ID="&downl_ID&"")
				SQLQueryNums=SQLQueryNums+1
				Response.Write("<br><font color='#FF0000'>已保存修改</font><br><br><a href='default.asp'>返回主页</a><br><br><a href='download.asp'>或者返回原下载<br><br>")
			End IF%>
          </td>
  </tr>   
</table>
      <%End IF%>
    </td>
  </tr>
</table>
<!--#include file="footer.asp" -->