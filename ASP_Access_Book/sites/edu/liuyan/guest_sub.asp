<% Sub Header(title)%>
<head>
<title>���ߴ���--<%=title%></title>
<mate content=text/html charset=gb2312 http-equiv=Content-Type>
<style type=text/css>
Body{font-size=9pt;font-famliy:����,Tahoma}
A:link {COLOR:red;font-size:9pt; TEXT-DECORATION: none}
A:visited {	COLOR:red; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:active {COLOR: #ff0000; TEXT-DECORATION: none}
A:hover {COLOR: blue;font-size:10pt; TEXT-DECORATION: underline}
TD {FONT-SIZE: 9pt}
</style>
<body bgcolor=#FFFFFF link=red>
<% End Sub%>

<% Sub Bottom() %>
<center><font color=#000000>
  Copyright &copy;���м���������������޹�˾ 2004
<br>
</center>
<% End Sub %>

<%sub redirect(url,success)%>
<html><head><title>OK</title><meta http-equiv=refresh content="0; url=<%=url%>"><body bgcolor="#ffffff" text="#000000" link="#000000" vlink="#000000" alink="#000000" topmargin=10>
<table width=100% height=100%><tr><td align=center><%=success%><br><a href="<%=url%>"><%=url%></a></font></td></tr></table></body></html>
<%end sub%>

<%SUB Msg(error) %>
<%Header "��ʾ"%>
<BR><BR>
<table border="0" cellpadding="20" cellspacing="0" width="85%" bgcolor=#ffffff align=center>
<tr><td><table cellspacing="0" cellpadding="2" width="100%" border="0"><tr bgcolor=#ffe1ff><td><font size=2 >��ʾ</td>
<td align=right><a href="javascript:history.go(-1)"><b>����</b></a></td></tr></table></td></tr>
<tr><td><%=error%></td></tr></table>
<p><p>
<%Bottom%>
</CENTER></font></BODY></HTML>
<%set guest=nothing
conn.Close
set conn = nothing
end sub%>

<%Sub tabRepl ()%>

	    <table border="1" align=center width="600" height="86" bordercolor="#000000" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF"> 
         <tr> 
          <td width="120" align=center height="86" rowspan="5" bgcolor="B5B5FF"> 
		   <font color=#000000>��! ��Һ�!</b><br>����:<%=rs("����")%></font><br> 
		   <IMG SRC="images/face/<%=rs("Ф��")%>.gif" ALT=<%=rs("����")%> BORDER="0"><br> 
           <%if rs("����")<>"" then%> 
		   ����:<%=rs("����")%></font> 
		   <%else%> 
		   ����:�����û</font> 
		   <%end if%>  
		 </td> 
         <td width="480" height="18" bgcolor="#FFE1FF"> 
		   <%if rs("����")="" then %> 
			&nbsp;<font color=#000000>����:&nbsp;&nbsp;<font color=BLUE>�˾�̫����,��Ը��д��!</font>  
		   <%else%>	 
			&nbsp;<font color=#000000>����:&nbsp;&nbsp;<font color=BLUE><%=rs("����")%></font> 
		   <%end if%> 
		  	<font color=#000000>&nbsp;&nbsp;&nbsp;����:</font><font color=BLUE><%=rs("��������")%></font> 
		  </td> 
       </tr> 
       <tr> 
         <td width="349" height="30" bgcolor="#FFFFFF"><%=rs("����")%></td> 
	   </tr> 
         <tr> 
         <td width="349" height="22" bgcolor="FFE1FF"><font color=#FFFFFF>
		 <%if rs("�ظ�����")="" then %> 
            &nbsp;<font color=000000>�ظ�����:&nbsp;&nbsp;�ޱ���</font>  
		   <%else%>	 
			&nbsp;<font color=000000>�ظ�����:&nbsp;&nbsp;<font color=BLUE><%=rs("�ظ�����")%></font> 
		   <%end if%> 
		   &nbsp;&nbsp;&nbsp;<font color=000000>�ظ���:</font><font color=BLUE><%=rs("�ظ�����")%></font></td> 
         </tr> 
         <tr> 
         <td Valign=top width="450" height="30" bgcolor="#FFFFFF"><%=rs("�ظ�")%></td> 
         </tr> 
       <tr> 
         <td width="349" Valign=middle height="18" bgcolor="#FFE1FF">�� 
		   <%if rs("�ʼ�")>"" then %> 
		     <IMG SRC="images/mailto.gif" ALT=���ö� BORDER="0"><a href="mailto:<%=rs("�ʼ�")%>">����</a> 
           &nbsp; 
		   <%end if%> 
		   <%if rs("��ҳ")<>"" and rs("��ҳ")<>"http://" then%> 
		   <IMG SRC="images/hpurl.gif" ALT=HomePage BORDER="0"><a href=<%=rs("��ҳ")%>>��ҳ</a> 
		    &nbsp; 
		   <%end if%>  
			<IMG SRC="images/reply.gif" ALT=Reply BORDER="0"><a href=guest_reply.asp?PageNo=<%=PageNo%>&ID=<%=rs("ID")%>>�ظ�</a> 
		   &nbsp; 
		   <IMG SRC="images/dele.gif" ALT=Delete BORDER="0"><a href=guest_del.asp?PageNo=<%=PageNo%>&ID=<%=rs("ID")%>>ɾ��</a> 
		 </td> 
       </tr> 
      </table> 
	  <tr> 
	    <td align=right><font color=#OOOOOO>IP��ַ:</font><font color=BLUE><font color=#FFFFFF>[</font><%=rs("IP")%><font color=#FFFFFF> 
		]</font>&nbsp;&nbsp;<font color=#000000>ϵͳ:</font><font color=BLUE><font color=#FFFFFF>[</font><%=rs("ϵͳ")%><font color=#FFFFFF>]</font></font> 
          </font> 
	   </tr></table> 
<%End Sub%>

<%Sub tabMsg()%>
	    <table border="1" align=center width="600" height="86" bordercolor="#000000" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF"> 
         <tr> 
          <td width="120" align=center height="86" rowspan="5" bgcolor="B5B5FF"> 
		   <font color=#000000>��! ��Һ�!</b><br>����:<%=rs("����")%></font><br> 
		   <IMG SRC="images/face/<%=rs("Ф��")%>.gif" ALT=<%=rs("����")%> BORDER="0"><br> 
           <%if rs("����")<>"" then%> 
		   ����:<%=rs("����")%></font> 
		   <%else%> 
		   ����:�����û</font> 
		   <%end if%>  
		 </td> 
         <td width="480" height="18" bgcolor="#FFE1FF"> 
		   <%if rs("����")="" then %> 
			&nbsp;<font color=#000000>����:&nbsp;&nbsp;<font color=BLUE>�˾�̫����,��Ը��д��!</font>  
		   <%else%>	 
			&nbsp;<font color=#000000>����:&nbsp;&nbsp;<font color=BLUE><%=rs("����")%></font> 
		   <%end if%> 
		  	<font color=#000000>&nbsp;&nbsp;&nbsp;����:</font><font color=BLUE><%=rs("��������")%></font> 
		  </td> 
       </tr> 
       <tr> 
         <td width="450"  height="50" bgcolor="#FFFFFF"><%=rs("����")%></td> 
	   </tr> 

       <tr> 
         <td width="349" Valign=middle height="18" bgcolor="#FFE1FF">�� 
		   <%if rs("�ʼ�")>"" then %> 
		     <IMG SRC="images/mailto.gif" ALT=���ö� BORDER="0"><a href="mailto:<%=rs("�ʼ�")%>">����</a> 
           &nbsp; 
		   <%end if%> 
		   <%if rs("��ҳ")<>"" and rs("��ҳ")<>"http://" then%> 
		   <IMG SRC="images/hpurl.gif" ALT=HomePage BORDER="0"><a href=<%=rs("��ҳ")%>>��ҳ</a> 
		    &nbsp; 
		   <%end if%>  
			<IMG SRC="images/reply.gif" ALT=Reply BORDER="0"><a href=guest_reply.asp?PageNo=<%=PageNo%>&ID=<%=rs("ID")%>>�ظ�</a> 
		   &nbsp; 
		   <IMG SRC="images/dele.gif" ALT=Delete BORDER="0"><a href=guest_del.asp?PageNo=<%=PageNo%>&ID=<%=rs("ID")%>>ɾ��</a> 
		 </td> 
       </tr> 
      </table> 
	  <tr> 
	    <td align=right><font color=#OOOOOO>IP��ַ:</font><font color=BLUE><font color=#FFFFFF>[</font><%=rs("IP")%><font color=#FFFFFF> 
		]</font>&nbsp;&nbsp;<font color=#000000>ϵͳ:</font><font color=BLUE><font color=#FFFFFF>[</font><%=rs("ϵͳ")%><font color=#FFFFFF>]</font></font> 
          </font> 
	   </tr></table> 
<%End Sub%>
