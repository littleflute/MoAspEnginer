<!--#include file="conn.asp"-->
<!--#include file="inc/char.asp"-->
<%
   	dim sql
   	dim rs
	dim CurPage
	if request("id")="" then
		response.write "��û��ѡ�����ӰƬ���뷵��"
		response.end
	end if
	If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
		CurPage = 1
	Else
		CurPage = CINT(Request.QueryString("CurPage"))
	End If
	display = CurPage
	set rs=server.createobject("adodb.recordset")
	if request("action")="save" then
		grade=request("grade")
		if request("content")<>"" then
		content=replace(replace(replace(request("content"),"'","��"),"<","&lt;"),">","&gt;")
		end if
		downid=request("id")
		if session("truevote")<>downid then
			sql="select * from Dvote where (id is null)"
			rs.open sql,conn,1,3
			rs.addnew
			if request("content")<>"" then
			rs("content")=content
			end if
			rs("grade")=grade
			rs("downid")=downid
			session("truevote")=downid
			rs.update
			rs.close
		else
			response.write "<script>alert('�Բ������Ѿ�������ͶƱ��');window.close();</script>"
		end if
	end if
		sql="select * from Dvote where downid="&request("id")&" order by id desc"
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
			V_num=0
			pingrade=0
		else
			V_num=rs.recordcount
			pingrade=allgrade()/V_num
			function allgrade() 
			dim tmprs 
    			tmprs=conn.execute("Select sum(grade) As grades from Dvote where downid="&request("id")&"") 
    			allgrade=tmprs("grades") 
			set tmprs=nothing 
			if isnull(allgrade) then allgrade=0 
			end function 
		end if

%>
<HTML><HEAD><TITLE>����ͶƱ</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link rel="stylesheet" href="images/style.css">
</HEAD>
<BODY bgColor=#ffffff text=#008080>
<table border=0 cellpadding=0 cellspacing=1 width=100% align="center" bgcolor="#009999">
  <tbody> 
  <tr bgcolor="#FFFFFF"> 
    <td align=left height="16" colspan="2"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="16">
        <tr> 
          <td background="images/bk04.gif" height="16"> 
            <div align="center"></div>
          </td>
        </tr>
        <tr> 
          <td> 
            <p align="center">Ŀǰ��ӰƬ����<%=V_num%>�˴�֣��ۺ�����Ϊ<%=formatnumber(pingrade,1)%></p>
          </td>
        </tr>
        <tr> 
          <td align="center"> 
            <form method=post action="vote.asp">
               <input type="hidden" name="action" value="save">
                <input name=grade type=radio value=1>
                <font color="#FFFFFF"> <font color="#000000">1</font></font>  
                <input name=grade type=radio value=2> 
                <font color="#FFFFFF"> <font color="#000000">2</font></font>  
                <input checked name=grade type=radio value=3> 
                <font color="#000000">3</font>  
                <input name=grade type=radio value=4> 
                <font color="#000000">4</font>  
                <input name=grade type=radio value=5> 
                <font color="#000000">5</font>  
                <input name=id type=hidden value='<%=Request("id")%>'> 
                <br> 
                ������ۣ�  
                <input name=content type=text size="30" class="smallinput" maxlength="100"> 
                &nbsp;  
                <input type=submit value="�ύ" name="cmdok" class="buttonface"> 
            </form> 
			<hr width="90%" size="1" align="center"> 
          </td> 
        </tr> 
        <tr>  
          <td align="center">�ѷ�������</td> 
        </tr> 
        <tr>  
          <td>  
            <%if rs.eof and rs.bof then%> 
            ��ʱû������  
            <% 
else 
	'����ÿҳ��¼���� 
	RS.PageSize=10 
	Dim TotalPages 
	TotalPages = RS.PageCount 
 
	If CurPage>RS.Pagecount Then  
		CurPage=RS.Pagecount 
	end if 
 
	RS.AbsolutePage=CurPage 
	rs.CacheSize = RS.PageSize 
 
	Dim Totalcount 
	Totalcount =INT(RS.recordcount) 
 
	StartPageNum=1 
	do while StartPageNum+10<=CurPage 
		StartPageNum=StartPageNum+10 
	Loop 
     
	EndPageNum=StartPageNum+9 
 
	If EndPageNum>RS.Pagecount then EndPageNum=RS.Pagecount 
%>          </td> 
        </tr> 
        <tr>  
          <td>  
            <table board=0 width=100%> 
              <% 
	I=0 
	p=RS.PageSize*(Curpage-1) 
	do while (Not RS.Eof) and (I<RS.PageSize) 
	p=p+1 
%> 
              <tr>  
                <td width=90%> <font color="green">�� </font><%=rs("content")%>&nbsp;&nbsp;[��֣�<%=rs("grade")%>]</td> 
              </tr> 
              <% 
	I=I+1 
	RS.MoveNext 
	Loop 
%> 
            </table> 
          </td> 
        </tr> 
        <tr>  
          <td>  
            <table width="100%" border="0" cellspacing="0" cellpadding="3" align="center"> 
              <tr>  
                <td>ҳ�Σ� <font color="#CC0000"><%=CurPage%></font>/<%=TotalPages%></td> 
                <td align="right"> ҳ���� <a href="vote.asp?CurPage=<%=StartPageNum-1%>&id=<%=request("id")%>"><<</a>  
                  <% For I=StartPageNum to EndPageNum   
      if I<>CurPage then %> 
                  <a href="vote.asp?CurPage=<%=I%>&id=<%=request("id")%>"><%=I%></a>  
                  <% else %> 
                  <%=I%>  
                  <% end if %> 
                  <% Next %> 
                  <% if EndPageNum<RS.Pagecount then %> 
                  <a href="vote.asp?CurPage=<%=EndPageNum+1%>&id=<%=request("id")%>">����...</a>  
                  <%end if%> 
                  <b>|</b> <a href="vote.asp?id=<%=request("id")%>">ˢ ��</a>&nbsp;</td> 
              </tr><% 
end if 
%> 
            </table> 
          </td> 
        </tr> 
 
        <tr>  
          <td bgcolor="#FFFFFF">  
             
              <table width="100%" border="0" background="images/bk04.gif" cellspacing="0" cellpadding="0" height="16"> 
                <tr> 
                  <td></td> 
                </tr> 
              </table> 
             
          </td> 
        </tr> 
      </table> 
    </td> 
  </tr> 
  </tbody>  
</table> 
<CENTER> 
</CENTER> 
</BODY></HTML> 
