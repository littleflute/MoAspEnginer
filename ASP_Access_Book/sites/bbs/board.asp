<% Time1=Timer %>
<!--#include file="odbc_connection.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.asp"-->
<!--#include file="code.asp"-->
<html>
	<head>
		<title>С��̳</title>
<!-- #include file="an.htm" -->
<!-- #include file="menu.asp" -->
<!--#include file="login_include.asp"-->
<!--#include file="menu_include.asp"-->
<!--#include file="menulist_include.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="254"><a href="announce.asp?boardid=<%=boardid%>"><img src="pic/h4.gif" alt="��������"  title="��������" width="72" height="21" border="0"></a></td>
    <td width="378">&nbsp;</td>
    <td width="128"><img src="pic/team2.gif" width="20" height="20">������<%= banzhu %></td>
  </tr>
</table>
<table width="760" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
    <tr align="center" class="p9">       
    <td width="6%" background="pic/h3.gif">���</td>			
    <td width="55%" background="pic/h3.gif">����</td>
    <td width="6%" background="pic/h3.gif">�ظ�</td>
    <td width="6%" background="pic/h3.gif">���</td>
    <td width="15%" background="pic/h3.gif">������</td>
    <td width="12%" background="pic/h3.gif">����ʱ��</td>
		</tr>
		<%
		dim sql,rs,boardid
		boardid=Trim(Request.QueryString("boardid"))
		'��ΪҪ��ҳ��ʾ��ѯ��������������淽������һ��recordset����
		sql="select * from news where layer=1 and boardid="&boardid&" order by flag desc, sorttime desc"
		set rs=Server.CreateObject("ADODB.Recordset")
		rs.Open sql,db,1,1          
		if not rs.bof and not rs.eof then
			'������ҪΪ�˷�ҳ��ʾ
			dim page_size                         '����ÿҳ��������¼����
			dim page_no                           '���嵱ǰ�ǵڼ�ҳ����
			dim page_total                        '������ҳ������
			page_size=10                         'ÿҳ��ʾ10����¼
			if Request("page_no")="" then         '�����һ�δ򿪣���page_noΪ1�������ɴ� 
				page_no=1                         '�صĲ�������
			else
				page_no=cint(Request("page_no"))
			end if
			session("page_no")=page_no            '��page_no����session,�Ա�����ҳ����ʱ��
			rs.pagesize=page_size                 '����ÿҳ��������¼
			page_total=rs.pagecount               '������ҳ��
			rs.absolutepage=page_no               '���õ�ǰ��ʾ�ڼ�ҳ
			'����һ����ʾ��ǰҳ�����м�¼
			dim i,j
			i=0
			j=page_size                           '�ñ�������������ʾ��������¼
			do while not rs.eof and j>0            'ѭ��֪����ǰҳ�������ļ���β
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
      �ظ��� <%=ubbcode(nRS("title"))%></a></td>
					
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
    <td width="346"><a href="announce.asp?boardid=<%=boardid%>"><img src="pic/h4.gif" alt="��������"  title="��������" width="72" height="21" border="0"></a></td>
    <td width="414" class="p12"> 
      <%
	'�����ӳ���д���йظ�ҳ��������Ϣ
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