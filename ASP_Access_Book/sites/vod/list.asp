<!--#include file="conn.asp"-->
<%tmp = "http://" & request.servervariables("SERVER_NAME") & _            
left(request.servervariables("SCRIPT_NAME"),len(request.servervariables("SCRIPT_NAME"))-len("/list.asp"))
   	dim sql
   	dim rs
   	dim classname,classid,Nclassname,Nclassid
	dim lasthits
   	dim title
	if request("id")="" then
		response.write "��û��ѡ�����ӰƬ���뷵��"
		response.end
	end if
	set rs=server.createobject("adodb.recordset")
  	sql="select class.class,Nclass.Nclass,download.showname,download.bb,download.classid,download.Nclassid,download.lasthits from download,class,Nclass where download.classid=class.classid and download.Nclassid=Nclass.Nclassid and download.ID="&request("id")
 	rs.open sql,conn,1,1
 	if not rs.eof then
		showname=rs("showname")
		bb=rs("bb")
		classid=rs("classid")
		Nclassid=rs("Nclassid")
		classname=rs("class")
		Nclassname=rs("Nclass")
		lasthits=rs("lasthits")
 	end if
	rs.close
'����ÿ��ÿ������
	tdate=year(Now()) & "-" & month(Now()) & "-" & day(Now())
	if trim(lasthits)=trim(tdate) then
		sql="update download set dayhits=dayhits+1 where id="&request("id")
		conn.Execute(sql)
'		response.write "success"
	else
		sql="update download set dayhits=1 where id="&request("id")
		conn.Execute(sql)
'		response.write "error"
	end if
	sql="update download set hits=hits+1,lasthits='"&tdate&"' where ID="&request("id")
	conn.Execute(sql)

   	p_year=CInt(year(Now()))-CInt(year(lasthits))
   	p_month=CInt(month(Now()))-CInt(month(lasthits))
   	p_day=CInt(day(Now()))-CInt(day(lasthits))
   	period_time=((p_year*12+p_month)*30+p_day)
	if cint(period_time)=<cint(7) then
		sql="update download set weekhits=weekhits+1 where id="&request("id")
		conn.Execute(sql)
	else
		sql="update download set weekhits=1 where id="&request("id")
		conn.Execute(sql)
	end if
%>
<HTML><HEAD>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<TITLE><%=classname%>��<%=Nclassname%>��<%=showname%> <%=bb%></TITLE>

<link rel="stylesheet" href="images/style.css">

<script language="JavaScript">
<!--
function Showvote(id)
{
	var filename="vote.asp?id="+id;
	window.open(filename,"��ʾ����","scrollbars=yes,width=350,height=300,status=yes,resizable=yes");
}
//-->
</SCRIPT>
</HEAD>
<!--#include file="topMain.asp"-->
<table width="752" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#004C90">
  <tr> 
    <td width="24%" valign="top" bgcolor="#E8F0F8"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#004C90" height="40">
        <form method="post" name="myform" action="search.asp">
          <tr> 
            <td bgcolor="#E8F0F8" valign="middle" align="center"> ������ 
              <input type="text" name="keyword" class=textfield size=10  maxlength="50"  style="font-family: Arial">
              <input type="submit" name="Submit22" value="����" style="height='21'" class="botton">
            </td>
          </tr>
        </form>
      </table>
      <table border=0 cellpadding=0 cellspacing=0 width=100% align="center">
        <tr> 
          <td height="20" bgcolor="#004C90"> 
            <div align="center"><font color=#ffffff>��������Top10</font></div>
          </td>
        </tr>
        <tr align=middle> 
          <td bgcolor=#999999 valign="top"> 
            <table border=0 cellpadding=0 cellspacing=0 width=100% height="150">
              <tr><td height="20" bgcolor="#E8F0F8" width="4%"></td> 
                <td bgcolor=#E8F0F8 valign=top width="96%"><font style=line-height:150%> 
                  <%                                                                                  
	dim tdate                                                                                  
	tdate=year(Now()) & "-" & month(Now()) & "-" & day(Now())                                                                                  
	sql="select top 10 id,showname,bb from download where "                                                                                  
	sql=sql&" lasthits="&tdate&" and  dayhits>0 "                                                                                  
	sql=sql&" order by dayhits desc"                                                                                 
	
	rs.open sql,conn,1,1                                                                                  
	if rs.eof and rs.bof then                                                                                  
%>
                  ����û������ 
                  <%else%>
                  <%do while not rs.eof                                                                                     
response.write "<LI><A href=list.asp?id="&rs("id")&">"&rs("showname")&" "&rs("bb")&"</A></LI>"                                                                                     
i=i+1                                                   
if i>=10 then exit do                                                   
rs.movenext                                                 
loop                                                 
i="0"                                                 
	end if                                                 
	rs.close                                                                                         
%>
                  </font> </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <table border=0 cellpadding=0 cellspacing=0 width=100% align="center">
        <tr> 
          <td  height="20" bgcolor="#004C90"> 
            <div align="center"><font color=#5d91c0>`</font><font color=#ffffff>��������Top10</font></div>
          </td>
        </tr>
        <tr align=middle> 
          <td bgcolor=#999999 colspan=4 valign="top"> 
            <table border=0 cellpadding=0 cellspacing=0 width=100% height="150">
              <tr> <td height="20" bgcolor="#E8F0F8" width="4%"></td>
                <td bgcolor=#E8F0F8 valign=top width="96%"><font style=line-height:150%> 
                  <%                                                                                     
	OldWeek = WeekDay(Date())-1                                                                                     
	If OldWeek = 0 Then OldWeek = 7                                                                                     
	OldWeek = Date()-OldWeek                                                                                     
	NewWeek = Date()+(9-WeekDay(Date()))                                                                                     
	sql="select top 10 id,showname,bb from download where "                                                                                     
	sql=sql&"  (lasthits < " & NewWeek & ") And (lasthits > " & OldWeek & ") and weekhits>0 "                                                                                     
	sql=sql&" order by weekhits desc"                                                                                     
	rs.open sql,conn,1,1                                                                                     
	if rs.eof and rs.bof then                                                                                     
%>
                  ����û������ 
                  <%else%>
                  <%do while not rs.eof                                                                                     
response.write "<LI><A href=list.asp?id="&rs("id")&">"&rs("showname")&" "&rs("bb")&"</A></LI>"                                                                                     
i=i+1                                                
if i>=10 then exit do                                                
rs.movenext                                                
loop                                                
i="0"                                                
	end if                                                 
	rs.close                                                 
%>
                  </font> </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td width="76%" valign="top" bgcolor="#FFFFFF"> 
      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="5"></td>
        </tr>
        <tr>
          <td height="1" background="images/bj4.gif"><img src="images/spacer.gif" width="1" height="1"></td>
        </tr>
      </table>
      <table width="98%" border="0" align="center">
        <tr>
          <td bgcolor="#efefef">����λ�ã�<A href="index.asp">��ҳ</A> >> <A href='sort.asp?classid=<%=classid%>'><%=classname%></A> 
            >> <A href='sort.asp?classid=<%=classid%>&Nclassid=<%=Nclassid%>'><%=Nclassname%></A> 
            >> <%=showname%>&nbsp;<%=bb%></td>
        </tr>
      </table>
      <table border=0 cellpadding=0 cellspacing=0 width=98% align="center">
        <tr> 
          <td align=middle valign=top width=568 height="160" bgcolor="#FFFFFF"> 
            <%                                          
    sql="select * from download where id="&request("id")                                       
	rs.open sql,conn,1,1                                          
%>
            <table width="98%" border="0" cellspacing="1" cellpadding="3" bgcolor="#CCCCCC">
              <tr bgcolor="#FFFFFF"> 
                <td colspan="3"><b><%=rs("showname")%></b></td>
              </tr>
              <tr> 
                <td width="80" align="right" nowrap bgcolor="#FFFFFF">ӰƬ�汾��</td>
                <td width="231" bgcolor="#FFFFFF"><%=rs("bb")%></td>
                <td rowspan="8" align="center" width="215" bgcolor="#FFFFFF"> 
                  <%                                                                               
if not isnull(rs("images")) and rs("images")<> "" then                                                                               
			response.write "<a href=showpic.asp?show=big&id="&rs("id")&" target=_blank><img src=showpic.asp?id="&rs("id")&" border=0  width=133 height=100 alt=����Ŵ�></a>"                                                                               
			else                                                                               
			response.write "��ͼƬ"                                                                               
		end if                                                                               
%>
                </td>
              </tr>
              <tr> 
                <td width="80" align="right" nowrap bgcolor="#FFFFFF">ӰƬ���ͣ�</td>
                <td width="231" bgcolor="#FFFFFF"><%=Nclassname%></td>
              </tr>
              <tr> 
                <td width="80" align="right" nowrap bgcolor="#FFFFFF">���л�����</td>
                <td width="231" bgcolor="#FFFFFF"><%=rs("system")%></td>
              </tr>
              <tr> 
                <td width="80" align="right" nowrap bgcolor="#FFFFFF">��Ȩ��ʽ��</td>
                <td width="231" bgcolor="#FFFFFF"><%=rs("orders")%></td>
              </tr>
              <tr> 
                <td width="80" align="right" nowrap bgcolor="#FFFFFF">ӰƬ��С��</td>
                <td width="231" bgcolor="#FFFFFF"><%=rs("size")%></td>
              </tr>
              <tr> 
                <td width="80" align="right" nowrap bgcolor="#FFFFFF">ӰƬ���ۣ�</td>
                <td width="231" bgcolor="#FFFFFF"><img src="images/<%=rs("hot")%>star.gif"></td>
              </tr>
              <tr> 
                <td width="80" align="right" nowrap bgcolor="#FFFFFF">�������ڣ�</td>
                <td width="231" bgcolor="#FFFFFF"><%=rs("dateandtime")%></td>
              </tr>
              <tr> 
                <td width="80" align="right" nowrap bgcolor="#FFFFFF">������ӣ�</td>
                <td width="231" bgcolor="#FFFFFF"> 
                  <%if rs("fromurl")<>"" then%>
                  <a href="<%=rs("fromurl")%>">��ҳ</a> 
                  <%end if%>
                </td>
              </tr>
              <tr> 
                <td width="80" align="right" nowrap bgcolor="#FFFFFF">�������أ�</td>
                <td colspan="2" bgcolor="#FFFFFF"><%=rs("dayhits")%>&nbsp;&nbsp;���ܣ�<%=rs("weekhits")%>&nbsp;&nbsp;�ܼƣ�<%=rs("hits")%></td>
              </tr>
              <%
if session("user")="" and rs("club")="1" then
%>
              <tr> 
                <td width="80" align="right" nowrap bgcolor="#FFFFFF">���ص�ַ��</td>
                <td colspan="2" bgcolor="#FFFFFF"><font color="#FF0000">��Ϊ��ԱӰƬ����ע������ǵĻ�Ա</font></td>
              </tr>
              <%
		  else
if rs("filename")="" and rs("filename1")="" and rs("filename2")="" then
%>
              <tr> 
                <td width="80" height="19" align="right" nowrap bgcolor="#FFFFFF">���ص�ַ��</td>
                <td colspan="2" height="19" bgcolor="#FFFFFF">��ʱû������</td>
              </tr>
              <%
else

if rs("movie")<>"" then
if rs("movie")="win" then     
play="playwin.asp"     
pimg="images/win.gif"     
else     
play="playrm.asp"     
pimg="images/rm.gif"     
end if     
%>
              <script language="JavaScript">     
function windowOpen(loadpos)     
{controlWindow=window.open(loadpos,"surveywin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=610,height=400,status=yes,resizable=yes");     
}     
              </script>
              <tr> 
                <td width="80" height="19" align="right" nowrap bgcolor="#FFFFFF">���߲��ţ�</td>
                <td colspan="2" height="19" bgcolor="#FFFFFF"><img src="<%=pimg%>">&nbsp; 
                  <%if rs("filename")<>"" then%>
                  <a href=# onClick=windowOpen('<%=play%>?id=<%=rs("id")%>&downid=1')>[1]</a> 
                  <%end if%>
                  <%if rs("filename1")<>"" then%>
                  <a href=# onClick=windowOpen('<%=play%>?id=<%=rs("id")%>&downid=2')>[2]</a> 
                  <%end if%>
                  <%if rs("filename2")<>"" then%>
                  <a href=# onClick=windowOpen('<%=play%>?id=<%=rs("id")%>&downid=3')>[3]</a> 
                  <%end if%>
                </td>
              </tr>
              <%else%>
              <%if rs("filename")<>"" then%>
              <tr> 
                <td width="80" height="19" align="right" nowrap bgcolor="#FFFFFF">���ص�ַ1��</td>
                <td colspan="2" height="19" bgcolor="#FFFFFF"> <a href="download.asp?downid=1&id=<%=rs("id")%>" target="_blank"><%=tmp%>/download.asp?id=<%=rs("id")%>&downid=1</a> 
                </td>
              </tr>
              <%end if%>
              <%if rs("filename1")<>"" then%>
              <tr> 
                <td width="80" align="right" nowrap bgcolor="#FFFFFF">���ص�ַ2��</td>
                <td colspan="2" bgcolor="#FFFFFF"> <a href="download.asp?downid=2&id=<%=rs("id")%>" target="_blank"><%=tmp%>/download.asp?id=<%=rs("id")%>&downid=2</a> 
                </td>
              </tr>
              <%end if%>
              <%if rs("filename2")<>"" then%>
              <tr> 
                <td width="80" align="right" nowrap bgcolor="#FFFFFF">���ص�ַ3��</td>
                <td colspan="2" bgcolor="#FFFFFF"> <a href="download.asp?downid=3&id=<%=rs("id")%>" target="_blank"><%=tmp%>/download.asp?id=<%=rs("id")%>&downid=3</a> 
                </td>
              </tr>
              <%end if%>
              <%end if%>
              <%end if%>
              <%end if%>
              <tr> 
                <td width="80" align="right" nowrap valign="top" bgcolor="#FFFFFF">ӰƬ��飺</td>
                <td colspan="2" bgcolor="#FFFFFF"><%=rs("note")%></td>
              </tr>
              <tr> 
                <td height="14" width="80" align="center" nowrap bgcolor="#FFFFFF"> 
                  <p>��</p>
                </td>
                <td height="14" colspan="2" bgcolor="#FFFFFF"> 
                  <p align="right"> ��<a href="javascript:Showvote('<%=rs("id")%>')">����ͶƱ</a>��</p>
                </td>
              </tr>
              <tr> 
                <td colspan="3" bgcolor="#FFFFFF"><font color=red>��</font> ���߿���Ӱ��Ҫ��Щ���ߣ�<a class="date" href="search.asp?action=title&keyword=Windows Media Player" title="����鿴��ϸ����" target="_blank">Windows 
                  Media Player</a>��<a class="date" href="search.asp?action=title&keyword=RealPlayer" title="����鿴��ϸ����" target="_blank">RealPlayer</a>�� 
                </td>
              </tr>
              <tr> 
                <td bgcolor="#FFFFFF" width="80" nowrap> 
                  <p align="right">���ӰƬ�� 
                </td>
                <center>
                  <td colspan="2" bgcolor="#FFFFFF"> 
                    <%  set rs_xg=server.createobject("adodb.recordset")                                      
    sql_xg="select * from download where id<>"&rs("id")&" and showname like '%"&rs("showname")&"%' order by id desc"                                      
	rs_xg.open sql_xg,conn,1,1                     
	if rs_xg.eof and rs_xg.bof then                                      
response.write ("û�����ӰƬ")                                      
else                                      
do while not rs_xg.eof                                      
response.write "<a href=list.asp?id="&rs_xg("id")&" target=_blank>"&rs_xg("showname")&" "&rs_xg("bb")&"</A><br>"                                      
rs_xg.movenext                                      
loop                                    
end if                                    
rs_xg.close                                     
set rs_xg=nothing%>
                  </td>
                </center>
              </tr>
              <tr> 
                <td bgcolor="#FFFFFF" width="80" nowrap valign="top"> 
                  <p align="right">������ۣ� 
                </td>
                <td colspan="2" bgcolor="#FFFFFF"> 
                  <table border="0" width="100%" cellspacing="0" cellpadding="0">
                    <center>
                      <tr> 
                        <td width="100%"> 
                          <%  set rs_pl=server.createobject("adodb.recordset")                 
    sql_pl="select * from Dvote where downid="&rs("id")&" order by id desc"                        
	rs_pl.open sql_pl,conn,1,1                   
	if rs_pl.eof and rs_pl.bof then                       
response.write ("��ʱδ�жԴ�ӰƬ������!")                                      
else                                      
do while not rs_pl.eof                                      
response.write "<li>[��֣�"&rs_pl("grade")&"]&nbsp;<font color=green>"&rs_pl("content")&"</font></li>"                                      
i=i+1                                             
if i>=5 then exit do               
rs_pl.movenext               
loop               
i="0"                                   
end if                                    
rs_pl.close                                      
set rs_pl=nothing%>
                        </td>
                      </tr>
                    </center>
                    <tr> 
                      <td width="100%"> 
                        <p align="right"><a href="javascript:Showvote('<%=rs("id")%>')">����������� 
                          / ��������&gt;&gt;</a> 
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tbody> </tbody> 
      </table>
    </td>
  </tr>
</table>
<!--#include file="CopyRight.asp"-->                                                
<%rs.close                                       
set rs=nothing                                       
conn.close                                                                       
set conn=nothing%>                                                                       
</BODY></HTML>