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
   if trim(request.form("Submt"))<>"下一步"  then
   response.write"<script>alert('由于您非法操作,\n系统拒绝进行下一步.');history.back();</script>"
     errorcount=1
	 else
	 stat=trim(request.form("title"))
	 votecount=trim(request.form("votecount"))
	  end if
   if stat="" or votecount="" then  
    response.write"<script>alert('请按条件把资料填写详细');history.back();</script>"
      errorcount=1
	  else
	  if not isnumeric(votecount) then
	     response.write"<script>alert('投票选项的个数必需为数字\n第一个数字不能为零');history.back();</script>"
                errorcount=1
			   else 
				votecount=clng(votecount)
				if votecount=0 or left(votecount,1)=0 or votecount>9 then
		  response.write"<script>alert('投票选项总数必需是0-9的数字\n且第一个数字不能为0');history.back();</script>"
		        errorcount=1
				 end if
				end if
			end if
			if errorcount=0 then
			sql="select title from stat where title='"&stat&"'"
			    set rs=conn.execute(sql)
				end if
				if not rs.eof then
        response.write"<script>alert('对不起,此投票标题已经存在');history.back();</script>"
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
				response.write"<script>alert('发生未知错误,数据添加失败\n返回重试')</script>"
		        end if
		%>
		<form name="form1" method="post" action="votesave.asp?id=<%=id%>&allcount=<%=votecount%>">
  <table width="400" border="0" align="center" bgcolor="#EFEFEF">
    <tr bgcolor="#B1E1F3"> 
      <td colspan="2"> 
        <div align="center"><strong>添加投票选项</strong></div></td>
  </tr>
  
   <% dim i
    for i=1 to votecount
  %> <tr>
      <td width="28%">投票<%=i%></td>
    <td width="72%"><input type="text" name="vote<%=i%>"></td>
  </tr>
  <% next %>
  <tr> 
    <td>&nbsp;</td>
      <td> <div align="center">
          <input type="submit" name="Submit" value="下一步">
        </div></td>
  </tr>
</table>
</form>
<% end if %>
</body>
</html>
