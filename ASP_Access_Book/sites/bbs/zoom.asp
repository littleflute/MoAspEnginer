<!--#include file="odbc_connection.asp"-->
<!--#include file="function.asp"-->
<html>
	<head>
		<title>��̳С��̳</title>
		
<!-- #include file="an.htm" -->
<center>
  <table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td height="16" class="p9"> <p><a href="announce.asp">��������</a> | <a href="adduer.asp">�û�ע��</a> 
          | <a href="uerlist.asp">�û��б�</a>|</p></td>
    </tr>
    <tr>
      <td height="16" class="p9"> | 
          <% If session("AdminUID")=""  Then %>
          <a href="login.asp">��½</a> 
          <% Else %>
          <% =session("AdminUID") %>
          ��ã� | <a href="useredit.asp?id=<%=session("AdminUID")%>">�޸�����</a>| 
         <a href=javascript:openScript('duan.asp',420,320)><FONT COLOR=#000000>����Ϣ</FONT></a>| 
          <% If session("AdminUID")="kenvin" or session("AdminUID")="��������" or session("AdminUID")="51fei" or session("AdminUID")="cavie" Then%>
          <a href="oneedit.asp">������̳</a> 
          <% End If %>
          | <a href="logout.asp">�˳���½</a> 
          <% End If %>
          |
		</td>
    </tr>
  </table>
  <table width="760" border="1" align="center" cellpadding="0" cellspacing="0" bgcolor="#eeeeee" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
    <tr align="center" bgcolor="#cccccc" class="p9"> 
      <td width="7%">���</td>
			
      <td width="43%">����</td>
			
      <td width="8%">�ظ�</td>
			
      <td width="8%">���</td>
			
      <td width="8%">������</td>
			
      <td width="26%">����ʱ��</td>
		</tr>
		<%
		dim sql,rs
		'��ΪҪ��ҳ��ʾ��ѯ��������������淽������һ��recordset����
		sql="select * from news  order by bbs_id desc"
		set rs=Server.CreateObject("ADODB.Recordset")
		rs.Open sql,db,1            
		if not rs.bof and not rs.eof then
			'������ҪΪ�˷�ҳ��ʾ
			dim page_size                         '����ÿҳ��������¼����
			dim page_no                           '���嵱ǰ�ǵڼ�ҳ����
			dim page_total                        '������ҳ������
			page_size=10                          'ÿҳ��ʾ10����¼
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
	<a href="announce.asp">����������</a>&nbsp&nbsp&nbsp&nbsp
	<%
	'�����ӳ���д���йظ�ҳ��������Ϣ
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