<% Response.Buffer=True %>
<!--#include file="../inc/admin.inc"-->
<% if request("del")<>"" then 
   conn.Execute("delete  from jobnews where id="&request("del"))
   end if %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" href="../inc/index.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�����˲š�&gt;�˲��г���&gt;���Ź���</title>
</head>
<SCRIPT LANGUAGE="JavaScript">
<!--//
function check()
{
   if (isNaN(go2to.page.value))
		alert("����ȷ��дת��ҳ����");
   else if (go2to.page.value=="") 
	     {
		alert("������ת��ҳ����");
		 }
   else
		go2to.submit();
}
//-->
</SCRIPT>
<SCRIPT language=JavaScript src="../inc/window.js"></SCRIPT>
<body topmargin="0" leftmargin="0">

  <table border="0" cellpadding="0" cellspacing="0" width="780" height="450">
    
  <tr bgcolor="#53BEB0"> 
    <td width="778" height="14" valign="top" colspan="3"></td>
    </tr>
  <tr>
    <td valign="top" height="424" width="176" bgcolor="#E1E1E1">
      ��
      <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" width="118" background="../images/stat-bg.GIF">
          <tr>
            <td height="34">
      <p align="center">
      <a href="mnews.asp">
      ���Ź���</a>
            </td>
          </tr>
          <tr>
            <td height="29">
      <p align="center"><a href="mperson.asp">��������û�</a>
            </td>
          </tr>
          <tr>
            <td height="34">
      <p align="center"><a href="mcompany.asp">����λ�û�</a>
            </td>
          </tr>
          <tr>
            <td height="34">
      <p align="center"><a href="#"onclick="javascript:if (confirm('�Ƿ�ȷ����յ�����������?')) href='reset.asp'; 
	    else return;">��յ�������</a>
            </td>
          </tr>
          <tr>
            <td height="28">
      <p align="center"><a href="exit.asp">�˳���¼</a>
            </td>
          </tr>
        </table>
        </center>
      </div>
      <p align="center">&nbsp; </td>
    <td valign="top" height="424" width="1" bgcolor="#53BEB0"> �� </td>  
	
	<td valign="top" height="424" width="602">
    <div align="center">
      <center>
     <% set rs=server.createobject("adodb.recordset") 
     if not isempty(request("page")) then   
     pagecount=cint(request("page"))   
     else   
     pagecount=1   
     end if
     sql="select * from jobnews order by id desc" 
	 rs.open sql,conn,1,1   
     if rs.eof and rs.bof then   
     response.write"<BR>" 
	 response.write"=== ���޼�¼ ==="
     response.write"<BR><BR>"
	 response.write"[<a href='addnews.asp' target='_blank'><font color='#000060'>�������|NEWS</font></a>]"
	 response.end  
	 end if
     rs.pagesize=10
     if pagecount>rs.pagecount or pagecount<=0 then              
     pagecount=1              
     end if              
     rs.AbsolutePage=pagecount %> 
	 <table border="1" cellpadding="0" cellspacing="0" width="545" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
        <tr>
          <td height="1" colspan="4" valign="bottom" width="541"><font color="#000000">&nbsp;</font></td> 
        </tr>
        <tr>
		   <% if rs.pagecount=1 then %>
		   <td width="541" height="6" colspan="4" valign="bottom">
          <font color="#000000">����[<font color="#ff0000"><%=rs.recordcount%></font>]������                        
		  ������[<font color="red">1��<%=rs.recordcount%></font>]��</font></td>             
        </tr>
        <tr>
          <td width="541" height="7" colspan="4" valign="bottom"></td>
        </tr>
        <%else%>
		<tr>
		 <td width="541" height="6" colspan="4" valign="bottom"><font color="#000000">
		 <% page_start=(pagecount-1)*rs.pagesize
            if pagecount=1 then page_start=1
		    page_end=rs.pagesize*pagecount
		    if pagecount*rs.pagesize=>rs.recordcount then page_end=rs.recordcount end if%>
         ����[<font color="#ff0000"><%=rs.recordcount%></font>]������                          
         ������[<font color="red"><%=page_start%>��<%=page_end%></font>]��</font></td>          
        </tr>
        <tr>
		  <td width="541" height="6" colspan="4" valign="bottom">
		  <% response.write"<form name=go2to form method=Post action='mnews.asp?key="+key+"'>"
		     response.write "<font color='#000064'>��&nbsp;</font>"                                              
		     if pagecount=1 then                                                        
			 response.write "<font color='#000064'>��ҳ ��һҳ</font>&nbsp;"
			 else                                                        
             response.write "<a href=mnews.asp?page=1&key="+key+"><font color='0000BE'>��ҳ</font></a>&nbsp;" 
	         response.write "<a href=mnews.asp?page="+cstr(pagecount-1)+"&key="+key+"><font color='0000BE'>��һҳ</font></a>&nbsp;"
			 end if                                             
             if rs.PageCount-pagecount<1 then                                                        
             response.write "<font color='#000064'>��һҳ βҳ</font>"                                                    
			 else                                                        
             response.write "<a href=mnews.asp?page="+cstr(pagecount+1)+"&key="+key+"><font color='0000BE'>��һҳ</font></a>&nbsp;"
			 response.write "<a href=mnews.asp?page="+cstr(rs.PageCount)+"&key="+key+"><font color='0000BE'>βҳ</font></a>"           
             end if 
			 response.write "<font color='000064'>&nbsp;ҳ��:<font color=blue>"&pagecount&"</font>/"&rs.pagecount&"ҳ</font>" 
			response.write "<font color='000064'> ת����<input type='text' name='page' size=2 maxLength=3 style='font-size: 9pt; color:#00006A; position: relative; height: 18' value="&PageCount&">ҳ</font>&nbsp;"                               
			response.write "<input class=button type='button' value='ȷ ��' onclick=check() style='font-family: ����; font-size: 9pt; color: #000073; position: relative; height: 19'>" %>         
             </td>
        </tr>
      </form>
	  <%end if%>
	  <tr>
      <td height="3" valign="top" colspan="4" bgcolor="#000000" width="541">
      </td>
        </tr>
      </center>
        <tr>
          <td width="45" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">NewsID</p>
          </td>
      <center>
          <td width="279" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;��   
            ��</td>
          <td height="18" bgcolor="#EBEEF3" valign="bottom" width="134"><p align="center">��������</p></td>
          <td height="18" bgcolor="#EBEEF3" valign="bottom" width="77">
            <p align="center">- ɾ�� -</td>     
        </tr>
         <% do while not rs.eof %> 
		<tr>
      <td width="45" height="18" bgcolor="#EBEEF3" valign="bottom">
        <p align="center"><%=rs("id")%></p>
      </td>
  <td width="279" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<a href="javascript:openwin('../viewnews.asp?id=<%=rs("id")%>','top=20,left=360,width=410,height=350')"><% if len(rs("title"))>20 then%><%=Left(rs("title"),20)%>...<%else%><%=rs("title")%><%end if%></a></td>
          <td width="134" height="18" bgcolor="#EBEEF3" valign="bottom"><p align="center">[<%=rs("idate")%>]</p></td>
          <td width="77" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center"><font color="#000046">[</font><a href="mnews.asp?del=<%=rs("id")%>"><font color="#000046">ɾ��</font></a><font color="#000046">]</font></td>
        </tr>
        <% i=i+1                                                                                                  
          rs.movenext                                                                                                  
          if i>=rs.PageSize then exit do 
		  loop                                                                    
          rs.close                                                                                                
          set rs=nothing                                                                                                
          conn.close                                                                                                
          set conn=nothing %>
		<tr>
      <td height="3" valign="top" colspan="4" bgcolor="#000000" width="541">
      </td>
        </tr>
      </center>
        <tr>
          <td height="8" valign="bottom" colspan="4" width="541">
            <p align="center"><br>
      <center>
      [<a href="javascript:openwin('addnews.asp','top=20,left=200,width=500,height=410')"><font color="#000060">�������|NEWS</font></a>]</center>
          </td>
        </tr>
      </table>
    </div>
    </td>
  </tr>
    
  <tr bgcolor="#53BEB0"> 
    <td valign="top" height="12" width="780" colspan="3"> �� </td>
    </tr>
  </table>
</body>

</html>