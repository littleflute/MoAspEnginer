<% Response.Buffer=True %>
<!--#include file="odbc_connection.asp"-->
<!--#include file="check.asp"-->
<!--#include file="char.asp"-->
<html>
<head>
	<title>����������</title>
<!-- #include file="an.htm" -->
���������� 
<center>
  <table border="1" align="center" cellpadding="0" cellspacing="0" bgcolor="#eeeeee" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" width=90%>
    <form method="post" action="" name="myform" >
      <tr> 
        <td class="p9">���⣺</td>
        <td><input type="text" name="title" size="40">
          ** </td>
      </tr>
            <tr> 
        <td class="p9">ͷ��</td>
        <td><img src="images/20.gif" alt=����������� align="middle" class=t2 id=idface>        <select                                  
      class=t2                                 
      onChange="document.images['idface'].src=options[selectedIndex].value;"                                 
      size=1 name=face>                                
          <option value=images/01.gif>�û�ͷ��-01                
          <option                
        value=images/02.gif>�û�ͷ��-02                
          <option value=images/03.gif>�û�ͷ��-03                
          <option                
        value=images/04.gif>�û�ͷ��-04                
          <option value=images/05.gif>�û�ͷ��-05                
          <option                
        value=images/06.gif>�û�ͷ��-06                
          <option value=images/07.gif>�û�ͷ��-07                
          <option                
        value=images/08.gif>�û�ͷ��-08                
          <option value=images/09.gif>�û�ͷ��-09                
          <option                
        value=images/10.gif>�û�ͷ��-10                
          <option value=images/11.gif>�û�ͷ��-11                
          <option                
        value=images/12.gif>�û�ͷ��-12                
          <option value=images/13.gif>�û�ͷ��-13                
          <option                
        value=images/14.gif>�û�ͷ��-14                
          <option value=images/15.gif>�û�ͷ��-15                
          <option                
        value=images/16.gif>�û�ͷ��-16                
          <option value=images/17.gif>�û�ͷ��-17                
          <option                
        value=images/18.gif>�û�ͷ��-18                
          <option value=images/19.gif>�û�ͷ��-19                
          <option                
        value=images/20.gif selected>�û�ͷ��-20</option>               
        </select><font class=j1>[<a 
      href="tx.htm" target=_blank>����ͷ��</a>]</font>
           </td>
      </tr>
      <tr> 
        <td class="p9">���ݣ�</td>
        <td><textarea name="body" rows="4" cols="40" wrap="soft"></textarea> <input name="user_name" type="hidden" id="user_name" value="<%=session("AdminUID")%>"></td>
      </tr>
      <tr>
        <td></td>
        <td><input type="submit" value="���ύ��" size="20" style="font-family: ����; font-size: 9pt; background-color: #FFFFFF; border-style: solid; border-width: 1"></td>
      </tr>
    </form>
  </table>
	</center>
	
<p align=center class="p9"><a href="index.asp?page_no=<%=session("page_no")%>"> 
  ������ҳ</a> 
  <%
	if request("title")<>"" and request("user_name")<>"" then
		dim title,body,layer,parent_id,child,hits,ip,user_name     '�����������ʹ��
		title=request.form("title")                                '�������±���
		body=ubbcode(request.form("body"))                                 '������������
		user_name=request.form("user_name")                        '������������
		layer=1                                                    '���ǵ�һ��
		parent_id=0                                                '��Ϊ�ǵ�һ�㣬�������Ϊ0
		child=0                                                    '�ظ�������ĿΪ0
		hits=0                                                     '�����Ϊ0
		face=request.form("face")
		ip=Request.ServerVariables("remote_addr")                  '����IP��ַ
		'���½����±��浽���ݿ�
		dim sql,svalues
		SQL = "Insert into news(title,layer,parent_id,child,hits,ip,user_name,submit_date,face"
		svalues = "values('" & title & "'," & layer & "," & parent_id & "," &child & "," & hits & ",'" & ip & "','" & user_name & "','" & date() & "','" & face & "'"
		if body<>"" then                                           '��������ݣ������body�ֶ�
			sql = sql & ",body"
			svalues = svalues & "," & "'" & body & "'"
		end if
		sql = sql & ") " & svalues & ")"
		db.execute(sql)
		db.close                                                    '�ر�connection����
		'������ϣ��ض������ҳ
		response.redirect "index.asp?page_no=" & session("page_no") 
	end if
	%>
</body>
</html>










