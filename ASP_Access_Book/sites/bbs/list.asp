<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="odbc_connection.asp"-->
<!--#include file="char.asp"-->
<!--#include file="code.asp"-->
<html>
<head>
<title>�û��б�</title>
<!-- #include file="an.htm" -->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid; background-color: #e3f1d1;">
  <tr> 
    <td width="475" height="24"><a href="index.asp">��̳��̳</a>��<% if Trim(Request.QueryString("order"))="2" then %>
      <a href="list.asp?order=2">���Ż���</a>
<% Else %>
      <a href="list.asp?order=1">��̳����</a>
<% End If %>
</td>
    <td width="283"> <a href="list.asp?order=1">��̳����</a> | <a href="list.asp?order=2">���Ż���</a> 
      | <a href="uerlist.asp?order=2">��������</a> | <a href="uerlist.asp">�û��б�</a> |</td>
  </tr>
    <tr align="center"> 
    <td height="15" colspan="9" background="pic/h3.gif" class="p9"><strong><font color="#FFFFFF"><% if Trim(Request.QueryString("order"))="2" then %>���Ż���  <% Else %>
��̳����<% End If %></font></strong></td>
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
if Trim(Request.QueryString("order"))="2" then
sql = "SELECT * FROM news where layer=1 order by child desc,sorttime desc"
else
sql = "SELECT * FROM news where layer=1 order by sorttime desc"
end if
		dim sql,rs
		'��ΪҪ��ҳ��ʾ��ѯ��������������淽������һ��recordset����
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
  <tr align="center" bgcolor="#FFFFFF"> 
    <td><%=(page_no-1)*page_size+i%> 
    <td align="left" class="p9"><a href="count_hits.asp?bbs_id=<%=rs("bbs_ID")%>&boardid=<%=rs("boardid") %>"><%=ubbcode(RS("title"))%></a></td>
    <td class="p9"><%=RS("child")%></td>
    <td class="p9"><%=RS("hits")%></td>
    <td class="p9"><a href="uerparticle.asp?id=<%=server.HTMLEncode(rs("user_name"))%>"><%=rs("user_name")%></a></td>
    <td class="p9"><% sub_time=rs("submit_date")
	   Response.Write year(sub_time)&"-"&month(sub_time)&"-"&day(sub_time) %></td>
  </tr>
  <% 
		   rs.movenext
		   loop
		end if
		rs.close
		%>
</table>
<%
Set rs = Nothing
db.close
set db=nothing
%>
<!--#include file="foot.asp"-->

