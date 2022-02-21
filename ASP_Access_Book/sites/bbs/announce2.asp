<% Response.Buffer=True %>
<!--#include file="odbc_connection.asp"-->
<!--#include file="check.asp"-->
<!--#include file="char.asp"-->
<html>
<head>
	<title>发表新文章</title>
<!-- #include file="an.htm" -->
发表新文章 
<center>
  <table border="1" align="center" cellpadding="0" cellspacing="0" bgcolor="#eeeeee" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" width=90%>
    <form method="post" action="" name="myform" >
      <tr> 
        <td class="p9">主题：</td>
        <td><input type="text" name="title" size="40">
          ** </td>
      </tr>
            <tr> 
        <td class="p9">头像：</td>
        <td><img src="images/20.gif" alt=个人形象代表 align="middle" class=t2 id=idface>        <select                                  
      class=t2                                 
      onChange="document.images['idface'].src=options[selectedIndex].value;"                                 
      size=1 name=face>                                
          <option value=images/01.gif>用户头像-01                
          <option                
        value=images/02.gif>用户头像-02                
          <option value=images/03.gif>用户头像-03                
          <option                
        value=images/04.gif>用户头像-04                
          <option value=images/05.gif>用户头像-05                
          <option                
        value=images/06.gif>用户头像-06                
          <option value=images/07.gif>用户头像-07                
          <option                
        value=images/08.gif>用户头像-08                
          <option value=images/09.gif>用户头像-09                
          <option                
        value=images/10.gif>用户头像-10                
          <option value=images/11.gif>用户头像-11                
          <option                
        value=images/12.gif>用户头像-12                
          <option value=images/13.gif>用户头像-13                
          <option                
        value=images/14.gif>用户头像-14                
          <option value=images/15.gif>用户头像-15                
          <option                
        value=images/16.gif>用户头像-16                
          <option value=images/17.gif>用户头像-17                
          <option                
        value=images/18.gif>用户头像-18                
          <option value=images/19.gif>用户头像-19                
          <option                
        value=images/20.gif selected>用户头像-20</option>               
        </select><font class=j1>[<a 
      href="tx.htm" target=_blank>所有头像</a>]</font>
           </td>
      </tr>
      <tr> 
        <td class="p9">内容：</td>
        <td><textarea name="body" rows="4" cols="40" wrap="soft"></textarea> <input name="user_name" type="hidden" id="user_name" value="<%=session("AdminUID")%>"></td>
      </tr>
      <tr>
        <td></td>
        <td><input type="submit" value="　提交　" size="20" style="font-family: 宋体; font-size: 9pt; background-color: #FFFFFF; border-style: solid; border-width: 1"></td>
      </tr>
    </form>
  </table>
	</center>
	
<p align=center class="p9"><a href="index.asp?page_no=<%=session("page_no")%>"> 
  返回首页</a> 
  <%
	if request("title")<>"" and request("user_name")<>"" then
		dim title,body,layer,parent_id,child,hits,ip,user_name     '定义变量方便使用
		title=request.form("title")                                '返回文章标题
		body=ubbcode(request.form("body"))                                 '返回文章内容
		user_name=request.form("user_name")                        '返回作者姓名
		layer=1                                                    '这是第一层
		parent_id=0                                                '因为是第一层，父编号设为0
		child=0                                                    '回复文章数目为0
		hits=0                                                     '点击数为0
		face=request.form("face")
		ip=Request.ServerVariables("remote_addr")                  '作者IP地址
		'以下将文章保存到数据库
		dim sql,svalues
		SQL = "Insert into news(title,layer,parent_id,child,hits,ip,user_name,submit_date,face"
		svalues = "values('" & title & "'," & layer & "," & parent_id & "," &child & "," & hits & ",'" & ip & "','" & user_name & "','" & date() & "','" & face & "'"
		if body<>"" then                                           '如果有内容，则添加body字段
			sql = sql & ",body"
			svalues = svalues & "," & "'" & body & "'"
		end if
		sql = sql & ") " & svalues & ")"
		db.execute(sql)
		db.close                                                    '关闭connection对象
		'保存完毕，重定向回首页
		response.redirect "index.asp?page_no=" & session("page_no") 
	end if
	%>
</body>
</html>










