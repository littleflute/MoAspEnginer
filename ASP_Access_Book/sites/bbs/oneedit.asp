<!--#include file="odbc_connection.asp"-->
<!--#include file="function2.asp"-->
<!--#include file="check.asp"-->
<html>
	<head>
		<title>С��̳</title>
		
<!-- #include file="an.htm" -->
<!-- #include file="menu.asp" -->
<% If (not session("AdminUID")="") and session("Adminflag")="0" Then %>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid; background-color: #e3f1d1;">
    <td width="469" height="16"><a href="index.asp">��̳</a> ��<a href="oneedit.asp">��������</a> 
      ���� <a href="sql.asp">ִ��sql���</a> | </td>
    <td width="289">����˵���| <a href="board.asp?boardid=0">�鿴����</a> <a href="oneedit.asp?pub=yes">| 
      ������ </a> | <a href="announce.asp?boardid=0">�������� </a> |</td>
  </tr>
</table>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="27" class="p9"> <p> 
        <% =session("AdminUID") %>
        ��<font color="#FF0000">�����ע�⣡�������Ҫɾ����һ��Ҫ����������ڻ������޸ģ�ֻ��ɾ����ɾ���󽫲���ȡ��������<br>
        ��: ���ر�ע�⣺���ɾ���������⣬��ô�����������лظ�������ɾ����������ǻظ������£���ֻɾ���˻ظ������ⲻ��Ӱ��</font></p></td>
  </tr>
</table>
<table width="760" border="0" align="center" cellspacing="2" bordercolor="#8800FF" >
  <tr align="center" class="p9"> 
    <td width="5%"  background="pic/h3.gif">���</td>
    <td  background="pic/h3.gif">����</td>
    <td width="10%" background="pic/h3.gif">����</td>
    <td width="5%" background="pic/h3.gif">�ظ�</td>
    <td width="9%" background="pic/h3.gif">������</td>
    <td width="15%" background="pic/h3.gif">����ʱ��</td>
    <td width="11%" background="pic/h3.gif">IP</td>
    <td width="5%" background="pic/h3.gif">ɾ��</td>
  </tr>
  <%
		dim sql,rs
		'��ΪҪ��ҳ��ʾ��ѯ��������������淽������һ��recordset����
	    if request("pub")="yes" then
		sql="select * from news where boardid= 0 order by submit_date desc"
		else	
		sql="select * from news  order by submit_date desc"
		end if
		set rs=Server.CreateObject("ADODB.Recordset")
		rs.Open sql,db,1,1            
		if not rs.bof and not rs.eof then
			'������ҪΪ�˷�ҳ��ʾ
			dim page_size                         '����ÿҳ��������¼����
			dim page_no                           '���嵱ǰ�ǵڼ�ҳ����
			dim page_total                        '������ҳ������
			page_size=20                          'ÿҳ��ʾ10����¼
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
			do while not rs.eof and j>0           'ѭ��֪����ǰҳ�������ļ���β
				i=i+1
				j=j-1
				board_id=RS("boardid")
		%>
  <tr align="center"> 
    <td height="20" bgcolor="#FFFFCC"><%=RS("bbs_id")%> 
    <td bgcolor="#FFFFCC" class="p9"><a href="count_hits.asp?bbs_id=<%=rs("bbs_ID")%>&boardid=<%= board_id %>"><%=RS("title")%></a></td>
    <td bgcolor="#FFFFCC" class="p9"> <%
sqlboardid="select * from board where boardid="&board_id
set rsboardid= db.execute(sqlboardid)
%> <a href="board.asp?boardid=<%=board_id%>"><%= rsboardid("name") %></a> <%
rsboardid.close
set rsboardid=nothing
%> </td>
    <td bgcolor="#FFFFCC" class="p9"><%=RS("child")%></td>
    <td bgcolor="#FFFFCC" class="p9"><a href="uerparticle.asp?id=<%=rs("user_name")%>"><%=rs("user_name")%></a></td>
    <td bgcolor="#FFFFCC" class="p9"><%=rs("submit_date")%></td>
    <td bgcolor="#FFFFCC" class="p9"><%=rs("ip")%></td>
    <td bgcolor="#FFFFCC" class="p9"><a href="del.asp?bbs_id=<%=RS("bbs_id")%>"><font color="#FF0000">ɾ��</font></a></td>
  </tr>
  <%
				rs.movenext
			loop
		end if
		%>
</table>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center"> 
      <%
	'�����ӳ���д���йظ�ҳ��������Ϣ
	call select_page(page_no,page_total)
	%> </td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
</table>

<% Else %>

<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="center">&nbsp;��û�й���Ȩ�޻������µ�½��<a href="javascript:history.go(-1)">�������ﷵ�أ�</a> 
	</td>
  </tr>
</table>
<% End If %>
<!--#include file="foot.asp"-->
