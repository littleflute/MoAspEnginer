<!--#include file="rscoon_1.asp"-->
<!--#include file="isadmin.asp"-->
<html>
<head>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../xsmz.css" rel=stylesheet>
</head>
<body>
<% dim stat,votecount,errorcount
   errorcount=0
   if trim(request.form("Submt"))<>"��һ��"  then
   response.write"<script>alert('�������Ƿ�����,\nϵͳ�ܾ�������һ��.');history.back();</script>"
     errorcount=1
	 else
	 stat=trim(request.form("title"))
	 votecount=trim(request.form("votecount"))
	  end if
   if stat="" or votecount="" then  
    response.write"<script>alert('�밴������������д��ϸ');history.back();</script>"
      errorcount=1
	  else
	  if not isnumeric(votecount) then
	     response.write"<script>alert('ͶƱѡ��ĸ�������Ϊ����\n��һ�����ֲ���Ϊ��');history.back();</script>"
                errorcount=1
			   else 
				votecount=clng(votecount)
				if votecount=0 or left(votecount,1)=0 or votecount>9 then
		  response.write"<script>alert('ͶƱѡ������������0-9������\n�ҵ�һ�����ֲ���Ϊ0');history.back();</script>"
		        errorcount=1
				 end if
				end if
			end if
			if errorcount=0 then
			sql="select title from stat where title='"&stat&"'"
			    set rs=conn.execute(sql)
				end if
				if not rs.eof then
        response.write"<script>alert('�Բ���,��ͶƱ�����Ѿ�����');history.back();</script>"
                 errorcount=1
				  else
			lastip="0.0.0.0"
			sql="insert into stat(title,isvote,lastip) values('"&stat&"','false','"&lastip&"')"	
	             conn.execute(sql)
				 end if
           if errorcount=0 then
              sql="select id from stat where title='"&stat&"'"
			  set rs=conn.execute(sql)
			  if not rs.eof then
			  id=rs("id")
		        else
				response.write"<script>alert('����δ֪����,�������ʧ��\n��������')</script>"
		        end if
		%>
		<form name="form1" method="post" action="votesave.asp?id=<%=id%>&allcount=<%=votecount%>">
  <table width="400" border="0" align="center" bgcolor="#EFEFEF">
    <tr bgcolor="#B1E1F3"> 
      <td colspan="2"> 
        <div align="center"><strong>���ͶƱѡ��</strong></div></td>
  </tr>
  
   <% dim i
    for i=1 to votecount
  %> <tr>
      <td width="28%">ͶƱ<%=i%></td>
    <td width="72%"><input type="text" name="vote<%=i%>"></td>
  </tr>
  <% next %>
  <tr> 
    <td>&nbsp;</td>
      <td> <div align="center">
          <input type="submit" name="Submit" value="��һ��">
        </div></td>
  </tr>
</table>
</form>
<% end if %>
</body>
</html>
