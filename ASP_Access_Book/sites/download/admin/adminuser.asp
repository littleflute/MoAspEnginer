<%@ CODEPAGE = "936" %>


<%
  if session("admin")="" then
  response.redirect "admin.asp"
  else
	if session("flag")>1 then
		response.write "<br><p align=center>您没有操作的权限</p>"
		response.end
	end if
  end if

%>
<!--#include file="conn.asp"-->
<HTML>
<HEAD>
<TITLE>用户管理</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<%
   Set rs=Server.CreateObject("Adodb.RecordSet")
   
   sql="select * from admin where flag>="&Session("flag")&" order by id"
   
   rs.Open sql,conn,1,1
%>
<BODY bgcolor="#39867B" leftmargin="0" topmargin="0">
<DIV align="center">
  <P>&nbsp;</P>
  
<P><FONT color="#000000">修改管理员信息</FONT> | <A href=adduser.asp>增加管理员</A></P>
  
	
<TABLE width="500" border="0" cellspacing="1" cellpadding="0">
		
<TR bgcolor="#E1F4EE"> 
      
<TD height="30" width="30"> 
        
<DIV align="center"><FONT color="#000000">序号</FONT></DIV>
      </TD>
      
<TD height="30" width="120"> 
        
<DIV align="center"><FONT color="#000000">用户名</FONT></DIV>
      </TD>
      
<TD height="30" width="120"> 
        
<DIV align="center"><FONT color="#000000">密码</FONT></DIV>
      </TD>
      
<TD height="30" width="100" bgcolor="#E1F4EE"> 
        
<DIV align="center"><FONT color="#000000">权限</FONT></DIV>
      </TD>
      
<TD height="30" width="130"> 
        
<DIV align="center"><FONT color="#000000">操作</FONT></DIV>
      </TD>
    </TR>
  </TABLE>
  <%while not rs.EOF %>
  <%if (rs("flag")>Session("flag")) or (rs("username")=Session("admin")) then%>
  <FORM method="post" action="saveuser.asp" style="margin:0">
    
		
<TABLE width="500" border="0" cellspacing="1" cellpadding="0">
			<TR bgcolor="#eeeeee"> 
				
<TD height="30" width="30" bgcolor="#E1F4EE"> 
					
<DIV align="center">1</DIV>
        </TD>
        
				
<TD height="30" width="120" bgcolor="#E1F4EE"> 
					<DIV align="center">
            <INPUT type="text" name="manager" value="<%=rs("username")%>" size="12">
          </DIV>
        </TD>
        
				
<TD height="30" width="120" bgcolor="#E1F4EE"> 
					<DIV align="center">
            <INPUT type="password" name="newpin" value="<%=rs("password")%>" size="12">
          </DIV>
        </TD>
        
				
<TD height="30" width="100" bgcolor="#E1F4EE"> 
					
<%if rs("flag")=1 then%>
					<DIV align="center">超级用户</DIV>
          <%end if%>
          <%if rs("flag")=2 then%> 
          <DIV align="center">普通用户</DIV>
          
<%end if%>
          
<%if rs("flag")=3 then%>
 
          <DIV align="center">员工</DIV>
          <%end if%>
        </TD>
        
				
<TD height="30" width="130" bgcolor="#E1F4EE"> 
					<DIV align="center">
            <INPUT type="submit" name="Submit" value="修改">&nbsp;
            <INPUT type="submit" name="Submit" value="删除">
            <INPUT type="hidden" name="oldmanager" value="<%=rs("username")%>">
            <INPUT type="hidden" name="oldpin" value="<%=rs("password")%>">
          </DIV>
        </TD>
      </TR>
    </TABLE>
  </FORM>
  <%else%>
    
	
<TABLE width="500" border="0" cellspacing="1" cellpadding="0">
		<TR bgcolor="#eeeeee"> 
			
<TD height="30" width="30" bgcolor="#E1F4EE"> 
				
<DIV align="center">1</DIV>
        </TD>
        
			
<TD height="30" width="120" bgcolor="#E1F4EE"> 
				
<DIV align="center">
            <%=rs("username")%>
          </DIV>
        </TD>
        
			
<TD height="30" width="120" bgcolor="#E1F4EEe"> 
				
<DIV align="center">
            ******
          </DIV>
        </TD>
        
			
<TD height="30" width="100" bgcolor="#E1F4EE"> 
				
<%if rs("flag")=1 then%>
				<DIV align="center">超级用户</DIV>
          
<%else%>
          <DIV align="center">普通用户</DIV>
          <%end if%>
        </TD>
        
			
<TD height="30" width="130" bgcolor="#E1F4EE"> 
				
<DIV align="center">&nbsp;
          </DIV>
        </TD>
      </TR>
    </TABLE>  
    
  <%
     end if
     
     rs.MoveNext
     
     Wend
  %>
  <P>&nbsp;</P> </DIV>
<%
  rs.Close
  set rs=Nothing
  
  conn.Close
  set conn=Nothing
%>
<P align=center><FONT color="#000000">超级用户可以进行所有功能的操作<BR>
普通用户没有用户管理和栏目管理权限，用户只有添加程序的权限。
</FONT>
</BODY>
</HTML>
