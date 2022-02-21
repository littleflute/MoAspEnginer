<!--#include file="odbc_connection.asp"-->
<!--#include file="function.asp"-->
<html>
	<head>
		<title>论坛小论坛</title>
		
<!-- #include file="an.htm" -->
<center>
  <table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td height="16" class="p9"> <p><a href="announce.asp">发表文章</a> | <a href="adduer.asp">用户注册</a> 
          | <a href="uerlist.asp">用户列表</a>|</p></td>
    </tr>
    <tr>
      <td height="16" class="p9"> | 
          <% If session("AdminUID")=""  Then %>
          <a href="login.asp">登陆</a> 
          <% Else %>
          <% =session("AdminUID") %>
          你好！ | <a href="useredit.asp?id=<%=session("AdminUID")%>">修改资料</a>| 
         <a href=javascript:openScript('duan.asp',420,320)><FONT COLOR=#000000>短消息</FONT></a>| 
          <% If session("AdminUID")="kenvin" or session("AdminUID")="汉穆拉比" or session("AdminUID")="51fei" or session("AdminUID")="cavie" Then%>
          <a href="oneedit.asp">管理论坛</a> 
          <% End If %>
          | <a href="logout.asp">退出登陆</a> 
          <% End If %>
          |
		</td>
    </tr>
  </table>
  <table width="760" border="1" align="center" cellpadding="0" cellspacing="0" bgcolor="#eeeeee" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
    <tr align="center" bgcolor="#cccccc" class="p9"> 
      <td width="7%">序号</td>
			
      <td width="43%">主题</td>
			
      <td width="8%">回复</td>
			
      <td width="8%">点击</td>
			
      <td width="8%">发言人</td>
			
      <td width="26%">发言时间</td>
		</tr>
		<%
		dim sql,rs
		'因为要分页显示查询结果，所以用下面方法创建一个recordset对象
		sql="select * from news  order by bbs_id desc"
		set rs=Server.CreateObject("ADODB.Recordset")
		rs.Open sql,db,1            
		if not rs.bof and not rs.eof then
			'以下主要为了分页显示
			dim page_size                         '定义每页多少条记录变量
			dim page_no                           '定义当前是第几页变量
			dim page_total                        '定义总页数变量
			page_size=10                          '每页显示10条记录
			if Request("page_no")="" then         '如果第一次打开，则page_no为1，否则，由传 
				page_no=1                         '回的参数决定
			else
				page_no=cint(Request("page_no"))
			end if
			session("page_no")=page_no            '将page_no存入session,以备其它页返回时用
			rs.pagesize=page_size                 '设置每页多少条记录
			page_total=rs.pagecount               '返回总页数
			rs.absolutepage=page_no               '设置当前显示第几页
			'下面一段显示当前页的所有记录
			dim i,j
			i=0
			j=page_size                           '该变量用来控制显示多少条记录
			do while not rs.eof and j>0            '循环知道当前页结束或文件结尾
				i=i+1
				j=j-1
		%>
				<tr bgcolor="#eeeeee" align="center">
					<td><%=(page_no-1)*page_size+i%>
					
      <td class="p9"><a href="count_hits.asp?bbs_id=<%=rs("bbs_ID")%>"><%=RS("title")%></a></td>
					
      <td class="p9"><%=RS("child")%></td>
					
      <td class="p9"><%=RS("hits")%></td>
					
      <td class="p9"><a href="uerparticle.asp?id=<%=rs("user_name")%>"><%=rs("user_name")%></a></td>
					
      <td class="p9"><%=rs("submit_date")%></td>
				</tr>
		<%
				rs.movenext
			loop
		end if
		%>
	</table>
	<a href="announce.asp">发表新文章</a>&nbsp&nbsp&nbsp&nbsp
	<%
	'调用子程序，写出有关各页的链接信息
	call select_page(page_no,page_total)
	%>
	</center>
</body>
</html>
<script language="javascript">
function openUser(id) {
	var Win = window.open("dispuser.asp?name="+id,"openScript","width=350,height=300,resizable=1,scrollbars=1,menubar=0,status=1" );
}

function openScript(url, width, height){
	var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}

function openDis(bid,rid,id){
	self.location="dispbbs.asp?boardid="+bid+"&RootID="+rid+"&id="+id
}

function PopWindow()
{ 
openScript('messanger.asp?action=newmsg',420,320); 
}
var nn = !!document.layers;
var ie = !!document.all;

if (nn) {
  netscape.security.PrivilegeManager.enablePrivilege("UniversalSystemClipboardAccess");
  var fr=new java.awt.Frame();
  var Zwischenablage = fr.getToolkit().getSystemClipboard();
}

function submitonce(theform){
//if IE 4+ or NS 6+
if (document.all||document.getElementById){
//screen thru every element in the form, and hunt down "submit" and "reset"
for (i=0;i<theform.length;i++){
var tempobj=theform.elements[i]
if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
//disable em
tempobj.disabled=true
}
}
}
</script>