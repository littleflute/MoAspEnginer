<% Time1=Timer %>
<!--#include file="odbc_connection.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.asp"-->
<!--#include file="code.asp"-->
<html>
	<head>
		<title>小论坛</title>
<!-- #include file="an.htm" -->
<!-- #include file="menu.asp" -->
<!--#include file="login_include.asp"-->
<!--#include file="menu_include.asp"-->
<!--#include file="menulist_include.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="254"><a href="announce.asp?boardid=<%=boardid%>"><img src="pic/h4.gif" alt="发表文章"  title="发表文章" width="72" height="21" border="0"></a></td>
    <td width="378">&nbsp;</td>
    <td width="128"><img src="pic/team2.gif" width="20" height="20">版主：<%= banzhu %></td>
  </tr>
</table>
<table width="760" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
    <tr align="center" class="p9">       
    <td width="6%" background="pic/h3.gif">序号</td>			
    <td width="55%" background="pic/h3.gif">标题</td>
    <td width="6%" background="pic/h3.gif">回复</td>
    <td width="6%" background="pic/h3.gif">点击</td>
    <td width="15%" background="pic/h3.gif">发言人</td>
    <td width="12%" background="pic/h3.gif">发言时间</td>
		</tr>
		<%
		dim sql,rs,boardid
		boardid=Trim(Request.QueryString("boardid"))
		'因为要分页显示查询结果，所以用下面方法创建一个recordset对象
		sql="select * from news where layer=1 and boardid="&boardid&" order by flag desc, sorttime desc"
		set rs=Server.CreateObject("ADODB.Recordset")
		rs.Open sql,db,1,1          
		if not rs.bof and not rs.eof then
			'以下主要为了分页显示
			dim page_size                         '定义每页多少条记录变量
			dim page_no                           '定义当前是第几页变量
			dim page_total                        '定义总页数变量
			page_size=10                         '每页显示10条记录
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
				<tr align="center">
					<td bgcolor="#e3f1d1"><%=(page_no-1)*page_size+i%>
					
      <td align="left" bgcolor="#e3f1d1" class="p9"><a href="count_hits.asp?bbs_id=<%=rs("bbs_ID")%>&boardid=<%=boardid %>"><%=ubbcode(RS("title"))%></a></td>
					
      <td bgcolor="#e3f1d1" class="p9"><%=RS("child")%></td>
					
      <td bgcolor="#e3f1d1" class="p9"><%=RS("hits")%></td>
					
      <td bgcolor="#e3f1d1" class="p9"><a href="uerparticle.asp?id=<%=server.HTMLEncode(rs("user_name"))%>"><%=rs("user_name")%></a></td>
					
      <td bgcolor="#e3f1d1" class="p9"><% sub_time=rs("submit_date")
	   Response.Write year(sub_time)&"-"&month(sub_time)&"-"&day(sub_time) %></td>
				</tr>
	<% 
	nsql="select * from news where layer=2 and parent_id="&rs("bbs_ID")&"  order by bbs_id"
	set nrs=db.execute(nsql) 
	if  not nrs.eof then
	do while not nrs.eof or nrs.bof 
	
	%>
	<tr align="center">
					
      <td bgcolor="#ffffff">&nbsp; 
      
    <td align="left" bgcolor="#ffffff" class="p9"><a href="count_hits.asp?bbs_id=<%=nrs("bbs_ID")%>&boardid=<%=boardid %>"> 
      回复： <%=ubbcode(nRS("title"))%></a></td>
					
      <td bgcolor="#ffffff" class="p9">&nbsp;</td>
					
      <td bgcolor="#ffffff" class="p9"><%=nRS("hits")%></td>
					
      <td bgcolor="#ffffff" class="p9"><a href="uerparticle.asp?id=<%=server.HTMLEncode(nrs("user_name"))%>"><%=nrs("user_name")%></a></td>
					
      <td bgcolor="#ffffff" class="p9"><% sub_time=rs("submit_date")
	   Response.Write year(sub_time)&"-"&month(sub_time)&"-"&day(sub_time) %></td>
				</tr>			
		<% 
		  nrs.movenext
		  loop
		  end if
				rs.movenext
			loop
		end if
		
		rs.close
		%>
	</table>
	
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="346"><a href="announce.asp?boardid=<%=boardid%>"><img src="pic/h4.gif" alt="发表文章"  title="发表文章" width="72" height="21" border="0"></a></td>
    <td width="414" class="p12"> 
      <%
	'调用子程序，写出有关各页的链接信息
	call select_page(page_no,page_total)
	%>
    </td>
  </tr>
</table>
<% Time2=Timer %>
<!--#include file="foot.asp"-->

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