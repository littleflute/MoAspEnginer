<%
  if session("admin")="" then
  response.redirect "admin.asp"
  else
	if session("flag")>2 then
		response.write "<br><p align=center>��û�в�����Ȩ��</p>"
		response.end
	end if
  end if
%>
<!--#include file="conn.asp"-->
<!--#include file="../inc/const.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<title>�� �� �� �� �� Ƶ</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<%
dim rs
dim sql
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from Nclass order by Nclassid asc"
rs.open sql,conn,1,1
%>
<SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
        count = 0
        do while not rs.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("Nclass"))%>","<%= trim(rs("classid"))%>","<%= trim(rs("Nclassid"))%>");
        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.Nclassid.length = 0; 

    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.Nclassid.options[document.myform.Nclassid.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
</SCRIPT>
<body bgcolor=#468ea3>
<form method="POST"  name="myform"  action="adminsave.asp?id=<%=request("id")%>&action=edit">
  <TABLE width="602" border="0" align="center" cellspacing="1">
    <TR align="center"> 
      <TD colspan="4" height="30" bgcolor="#145f74" width="594"><b><FONT color="#FFFFFF">�� 
        �� �� �� �� Ƶ</FONT></b></TD>
    </TR>
    <tr> 
      <td width="75" align="right" valign="top" height="30" bgcolor="#a5d0dc"><b><font color="#FF0000">ӰƬ���ͣ�</font></b></td>
      <td width="225" bgcolor="#a5d0dc" height="30"> 
        <SELECT name="classid" onChange="changelocation(document.myform.classid.options[document.myform.classid.selectedIndex].value)" size="1">
          <%dim classid,Nclassid
	classid=request("classid")
	Nclassid=request("Nclassid")
  sql="select * from class"
 	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "Not record."
	else
	do while not rs.eof
        if classid=cstr(rs("classid")) then
               sel="selected"
        else
               sel=""
        end if	
	response.write "<option " & sel & " value='"+CStr(rs("classID"))+"' name=classid>"+rs("class")+"</option>"+chr(13)+chr(10)
	rs.movenext
    	loop
	end if
	rs.close
%>
        </SELECT>
      </td>
      <td width="75" bgcolor="#a5d0dc" align="right" height="30"> <B>ѡ����ࣺ</B> 
      </td>
      <td width="225" bgcolor="#a5d0dc" height="30"> 
        <SELECT name="Nclassid" size="1">
          <%             
  	sql="select * from Nclass"             
 	rs.open sql,conn,1,1             
	if rs.eof and rs.bof then             
		response.write "Not record."             
	else             
        do while not rs.eof             
        if Nclassid=cstr(rs("Nclassid")) then             
               sel="selected"             
        else             
               sel=""             
        end if	             
        response.write "<option " & sel & " value='" +  Cstr(rs("Nclassid")) + "'>" + rs("Nclass") + "</option>"             
        rs.MoveNext             
        Loop             
	end if             
        rs.close             
%>
        </SELECT>
      </td>
    </tr>
    <%             
	sql="select * from download where id="&request("id")             
	rs.open sql,conn,1,1             
%>
    <tr> 
      <td width="75" align="right" height="30" bgcolor="#CCE6ED"><b><font color="#FF0000">ӰƬ</font></b><font color="#FF0000"><b>���ƣ�</b></font></td>
      <td width="225" height="30" bgcolor="#CCE6ED"> <FONT color="#FF0000"> 
        <input type="text" name="txtshowname" size="33"            
          class="smallinput" maxlength="100" value="<%=rs("showname")%>">
        </FONT></td>
      <td width="75" height="30" bgcolor="#CCE6ED" align="right"> <b>ӰƬ�汾��</b></td>
      <td width="225" height="30" bgcolor="#CCE6ED"> <FONT color="#FF0000"> 
        <input type="text" name="txtbb" size="33"            
          class="smallinput" maxlength="100" value="<%=rs("bb")%>">
        </FONT> </td>
    </tr>
    <tr> 
      <td width="75" align="right" height="14" nowrap bgcolor="#A5D0DC">��Ӱ��ʽ:</td>
      <td width="450" height="14" colspan="3" bgcolor="#A5D0DC"> [<img border="0" src="../images/rm.gif"> 
        <input type="radio" value="rm" name="movie"<%if rs("movie")="rm" then%> checked<%end if%>>
        ]&nbsp;&nbsp;&nbsp;&nbsp; [<img border="0" src="../images/win.gif"> 
        <input type="radio" value="win" name="movie"<%if rs("movie")="win" then%> <%end if%> checked>
        ]&nbsp;&nbsp;&nbsp; [���ǵ�Ӱ 
        <input type="radio" value="" name="movie" >
        ]��������ߵ�Ӱ��ѡ�񲥷Ÿ�ʽ</td>
    </tr>
    <tr> 
      <td width="75" align="right" height="15" nowrap bgcolor="#A5D0DC">��Ա��</td>
      <td width="450" height="15" colspan="3" bgcolor="#A5D0DC">�Ƿ��Ա
<input type="checkbox" name="club" value="on" <%if rs("club")="1" then%> checked<%end if%>>&nbsp;���ѡ���Ա��ֻ�л�Ա�ſɹۿ�</td>
    </tr>
    <tr bgcolor="#cce6ed"> 
      <td width="75" align="right" height="30"><b><font color="#FF0000">���ص�ַ1</font>��</b></td>
      <td width="450" height="30" colspan="3"> 
        <input type="text" name="txtfilename" size="82"            
          class="smallinput" maxlength="100" value="<%=rs("filename")%>">
      </td>
    </tr>
    <tr> 
      <td width="75" align="right" height="30" bgcolor="#a5d0dc"><b>���ص�ַ2��</b></td>
      <td width="450" height="30" bgcolor="#a5d0dc" colspan="3"> 
        <input type="text" name="txtfilename1" size="82"             
          class="smallinput" maxlength="100" value="<%=rs("filename1")%>">
      </td>
    </tr>
    <tr bgcolor="#cce6ed"> 
      <td width="75" align="right" height="30"><b>���ص�ַ3��</b></td>
      <td width="450" height="30" colspan="3"> 
        <input type="text" name="txtfilename2" size="82"             
          class="smallinput" maxlength="100" value="<%=rs("filename2")%>">
      </td>
    </tr>
    <tr bgcolor="#cce6ed"> 
      <td width="75" align="right" height="30" bgcolor="#A5D0DC"><b>�Ƽ��ȣ�</b></td>
      <td width="450" height="30" colspan="3" bgcolor="#A5D0DC"> 
        <select name="hot" size="1">
          <option <%if rs("hot")=1 then%> selected <%end if%> value="1" name="hot">һ�Ǽ�</option>
          <option <%if rs("hot")=2 then%> selected <%end if%> value="2" name="hot">���Ǽ�</option>
          <option <%if rs("hot")=3 then%> selected <%end if%> value="3" name="hot">���Ǽ�</option>
          <option <%if rs("hot")=4 then%> selected <%end if%> value="4" name="hot">���Ǽ�</option>
          <option <%if rs("hot")=5 then%> selected <%end if%> value="5" name="hot">���Ǽ�</option>
        </select>
        <!--<input type="text" name="hot" size="60" class="smallinput" maxlength="100" value="3">-->
      </td>
    </tr>
    <tr> 
      <td width="75" align="right" height="30" bgcolor="#CCE6ED"><b>ӰƬ��С��</b></td>
      <td width="225" height="30" bgcolor="#CCE6ED"> 
        <!--<input type="text" name="system" size="60" class="smallinput" maxlength="100" value="Win98/2000+PWS4&NT+IIS4/5">-->
        <input type="text" name="size" size="33"             
          class="smallinput" maxlength="100" value="<%if rs("size")<>"" then%><%=rs("size")%><%else%>K<%end if%>">
      </td>
      <td width="75" height="30" bgcolor="#CCE6ED" align="right"> <b>���л�����</b> 
      </td>
      <td width="225" height="30" bgcolor="#CCE6ED"> 
        <select name="system" size="1">
          <option value="Win9x/WinNT/Win2000/WinME" selected name="system">Win9x/WinNT/Win2000/WinME</option>
          <option value="Win9x/WinNT/Win2000" name="system">Win9x/WinNT/Win2000</option>
          <option value="Windows9x/NT" name="system">Win9x/WinNT</option>
          <option value="Win9x/WinME" name="system">Win9x/WinME</option>
          <option value="WinNT/Win2000" name="system">WinNT/Win2000</option>
          <option value="ASP����" name="system">ASP����</option>
          <option value="PHP����" name="system">PHP����</option>
          <option value="CGI����" name="system">CGI����</option>
          <option value="JSP����" name="system">JSP����</option>
          <option value="Linux����" name="system">Linux����</option>
        </select>
      </td>
    </tr>
    <tr bgcolor="#cce6ed"> 
      <td width="75" align="right" height="30" bgcolor="#A5D0DC"><b>ӰƬ��Դ��</b></td>
      <td height="30" bgcolor="#A5D0DC" width="225"> 
        <input type="text" name="fromurl" size="33"             
          class="smallinput" maxlength="100" value="<%=rs("fromurl")%>">
      </td>
      <td height="30" bgcolor="#A5D0DC" width="75" align="right"><B>ӰƬͼƬ��</B></td>
      <td height="30" bgcolor="#A5D0DC" width="225"> 
        <INPUT type="text" name="images" size="33"             
          class="smallinput" maxlength="160" value="<%=rs("images")%>">
      </td>
    </tr>
    <tr bgcolor="#cce6ed"> 
      <td width="75" align="right" height="30"><b>��Ȩ��ʽ��</b></td>
      <td width="450" height="30" colspan="3"> 
        <select name="order" size="1">
          <option selected value="����ӰƬ" name="order">����ӰƬ</option>
          <option value="���ӰƬ" name="order">���ӰƬ</option>
          <option value="��ҵӰƬ" name="order">��ҵӰƬ</option>
        </select>
        <!--<input type="text" name="order" size="60" class="smallinput" maxlength="100" value="���">-->
        �Ƿ��Ƽ� 
        <input type="checkbox" name="hots" value="on" <%if rs("hots")="1" then%> checked<%end if%>>
        �Ƿ�Ʒ
<input type="checkbox" name="hide" value="on" <%if rs("stop")="1" then%> checked<%end if%>>
      </td>
    </tr>
    <tr bgcolor="#a5d0dc"> 
      <td width="75" align="right" valign="top" height="30"><b><font color="#FF0000">��ʽ���</font>��</b></td>
      <td width="450" colspan="3" height="30"> 
        <textarea rows="6" name="txtnote" cols="82" class="smallarea"><%
	content=replace(rs("note"),"<br>",chr(13))
	content=replace(content,"&nbsp;"," ")
	response.write content
%></textarea>
      </td>
    </tr>
  </table>             
                       
  <div align="center"><center><p><input type="submit" value=" �� �� "             
  name="cmdok" class="buttonface">&nbsp; <input type="reset" value=" �� �� "             
  name="cmdcancel" class="buttonface"></p>             
  </center></div>             
</form>             
</body>             
</html>             
<%             
	set rs=nothing             
	conn.close             
	set conn=nothing             
%>