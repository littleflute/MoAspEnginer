<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/library.asp" -->
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif" class="wordbreak">
  <tr>
    <td width="160" valign="top" bgcolor="#C8C7BD" nowrap><%
	Call MemberCenter
	Response.Write("<br>")
	Call SiteInfo
	Response.Write("<br>")
	Call NewBlogList
    Response.Write("<br>")
	Call NewCommList
	Response.Write("<br>")
	Call blogSearch%><br>
	
	</td>
    <td width="100%" valign="top" bgcolor="#FFFFFF">
	
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
          <tr>
            <td colspan="2" class="msg_head">Tags 搜索</td>
          </tr>
          <tr class="msg_content" bgcolor="#FFFFFF">
            <td><form action="bloglisttag.asp" method="get" name="SearchTag">
		   <input name="tags" type="text" size="36" maxlength="36"> <input type="submit" value=" 搜索 ">
		 </form></td>
          </tr>
          <tr class="msg_content" bgcolor="#FFFFFF">
            <td colspan="2"><b><font color="#FF0000">热门 Tags</font></b><BR><BR><%Call TagsList("Hot")%></td>
          </tr>
          <tr class="msg_content" bgcolor="#FFFFFF">
            <td colspan="2"><b><font color="#FF0000">Tags 列表</font></b><BR><BR><%Call TagsList("All")%></td>
          </tr>
      </table>
	  
	  </td>
  </tr>
</table>
<!--#include file="footer.asp" -->