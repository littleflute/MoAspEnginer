<%@ language=vbscript %>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="guest_sub.asp" -->
<% Set rs=server.createobject("ADODB.RECORDSET")
   sql="SELECT * FROM guest ORDER BY ID DESC"
   rs.open sql,conn,1,1
   if rs.bof and rs.eof then%> 
     <center></center>
	 <%Msg("<font color=#FF0000><li>�Բ���,�����κ�����!</font></li><br><center><a href='guest_input.asp'><<��������>></a></center>")
   else
     rs.pagesize=INT(Mypage)
	 PageNo=REQUEST("PageNo")
	 RS.AbsolutePage=PageNo
	 TSum=INT(rs.RECORDCOUNT/Mypage*-1)*-1
	 RowCount=rs.PageSize
	 if PageNo="" or PageNo=0 then
	   PageNo=1
	 else
	   PageNo=PageNo+1
	   PageNo=PageNo-1
     end if
	 if CINT(PageNo)>1 then
	    if CINT(PageNo)>CINT(TSum) then
		  Msg("<font size=2 color=#FF0000><li>�Բ���û������Ҫ��ҳ��</li></font>")
          Response.End
	    end if
	 end if		    
     if PageNo<0 then
	    Msg("<font color=#FF0000 size=2><li>û����һҳ!��������.</li><font>")
		Response.End
	 End if
%>
  <%Header(title)%> 
  <table border=0 cellspacing=0 align=center> 
    <tbody> 
	  <tr> 
	    
  <td align=center><img border="0" src="images/mini1.gif" width="723" height="88"></td> 
	   </tr> 
	   <tr> 
	     <td> 
  	      <table border=0 width=600 align=center> 
	        <td align=lift> 
            <font color="#0000FF">������:<%=rs.RECORDCOUNT%> /&nbsp;��ҳ��:<%=TSum%>&nbsp;   
            /&nbsp;Ŀǰ��:<%=PageNo%>ҳ</font>   
	       </td>   
	       <td align=right>   
		   <a href="guest_input.asp"><img src="images/write.gif" border=0>��д����</a>&nbsp;    
		   <a href=<%=home%>><img src="images/home.gif" border=0>������ҳ</a>    
	       </td></table></tr>    
	    <tr><td><HR SIZE=1 WIDTH=100% NOSHADE COLOR=#0000FF></td></tr>    
        <tr>    
		 <td align=left><font color="#0000FF"><%=PageList%></font><td></tr>   
		<tr><td>   
       <font color="#0000FF">   
	   <%DO WHILE NOT rs.EOF AND RowCount>0%>   
	   <br>   
	   <%if rs("�ظ�")<>"" then%>   
       </font>   
	    <table border=0 width=485 align=center>    
	   <tr><td>    
		<%tabRepl%>   
	       
	  <%else%>   
	   <table border=0 width=485 align=center>   
	   <tr><td>   
        <font color="#0000FF">   
		<%tabMsg   
	  end if   
      RowCount = RowCount - 1   
      rs.MoveNext    
      Loop     
	  set guest=nothing    
      set rs = nothing%></font></td></tr>   
	  <tr><td height=12><HR SIZE=1 WIDTH=100% NOSHADE COLOR=#0000FF></td></tr>   
	  <tr><td align=center><font color="#0FF0FF"><%Bottom%></font></tr></table>   
<%Function PageList()	   
 if TSum>1 then   
	if PageNo=1 then   
	   NextPage=PageNo+1   
	   response.write"<a href='guest_list.asp?PageNo="&NextPage&"'>��һҳ</a>"   
    end if   
	if PageNo=TSum then   
       PrwePage=PageNo-1    
	   response.write "<a href='guest_list.asp?PageNo="&PrwePage&"'>��һҳ</a>"   
	end if   
	if PageNo>1 and TSum>PageNo then   
	   NextPage=PageNo+1   
	   PrwePage=PageNo-1   
	   response.write "<a href='guest_list.asp?PageNo="&PrwePage&"'>��һҳ</a>&nbsp|"   
	   response.write"&nbsp;<a href='guest_list.asp?PageNo="&NextPage&"'>��һҳ</a>"   
	end if   
 end if   
End Function%>        
 <%end if%> 
 
 
 
