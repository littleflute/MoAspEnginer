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

<title>�� �� �� �� �� ��</title>
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
<body bgcolor=#FFFFFF>
<form method="POST"  name="myform"  action="adminsave.asp?id=<%=request("id")%>&action=edit">
  <TABLE width="610" border="0" cellspacing="1">
    <TR align="center" bgcolor="#E1F4EE"> 
      <TD colspan="4" height="30"><b> 
        <input type="submit" value=" �� �� "             
  name="cmdok2" class="buttonface">
        &nbsp; 
        <input type="reset" value=" �� �� "             
  name="cmdcancel2" class="buttonface">
        </b></TD>
    </TR>
    <tr> 
      <td width="72" align="right" valign="top" height="30" bgcolor="#E1F4EE"><b><font color="#FF0000">������ͣ�</font></b></td>
      <td width="250" bgcolor="#E1F4EE" height="30"> 
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
      <td width="67" bgcolor="#E1F4EE" align="right" height="30"> <B><font color="#FF0000">ѡ����ࣺ</font></B> 
      </td>
      <td width="208" bgcolor="#E1F4EE" height="30"> 
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
      <td width="72" align="right" height="26" bgcolor="#E1F4EE"><b><font color="#FF0000">���</font></b><font color="#FF0000"><b>���ƣ�</b></font></td>
      <td width="250" height="26" bgcolor="#E1F4EE"> <FONT color="#FF0000"> 
        <input type="text" name="txtshowname" size="33"            
          class="smallinput" maxlength="100" value="<%=rs("showname")%>">
        </FONT></td>
      <td width="67" height="26" bgcolor="#E1F4EE" align="right"> <b><font color="#FF0000">���л�����</font></b></td>
      <td width="208" height="26" bgcolor="#E1F4EE"> <FONT color="#FF0000"> 
        <select name="system" size="1">
          <option <%if rs("system")="Win9x/WinNT/Win2000/WinXP/WinME" then%> <%end if%> value="Win9x/WinNT/Win2000/WinXP/WinME" name="system"><font color="#FF0000">Win9x/WinNT/Win2000/WinXP/WinME</font></option>
          <option <%if rs("system")="Win9x/WinNT/Win2000/WinME" then%> <%end if%> value="Win9x/WinNT/Win2000/WinME" name="system"><font color="#FF0000">Win9x/WinNT/Win2000/WinME</font></option>
          <option <%if rs("system")="Win9x/WinNT/Win2000" then%> <%end if%> value="Win9x/WinNT/Win2000" name="system"><font color="#FF0000">Win9x/WinNT/Win2000</font></option>
          <option <%if rs("system")="Win9x/WinNT" then%> <%end if%> value="Win9x/WinNT" name="system"><font color="#FF0000">Win9x/WinNT</font></option>
          <option <%if rs("system")="Win9x/WinME" then%> <%end if%> value="Win9x/WinME" name="system"><font color="#FF0000">Win9x/WinME</font></option>
        </select>
        </FONT> </td>
    </tr>
    <tr> 
      <td width="72" align="right" height="30" nowrap bgcolor="#E1F4EE">��ʽ:</td>
      <td width="250" height="30" bgcolor="#E1F4EE"> [���
<input type="radio" value="" name="movie" checked>
        ]&nbsp;&nbsp;&nbsp; </td>
      <td width="67" height="30" bgcolor="#E1F4EE"><b><font color="#FF0000">�Ƽ��ȣ�</font></b></td>
      <td width="208" height="30" bgcolor="#E1F4EE"> 
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
    <tr bgcolor="#E1F4EE"> 
      <td width="72" align="right" height="30"> 
        <div align="left"><b><font color="#FF0000">���ص�ַ1</font>��
          </b></div>
      </td>
      <td width="250" height="30"> 
        <input type="text" name="txtfilename" size="35"            
          class="smallinput" maxlength="100" value="<%=rs("filename")%>">
      </td>
      <td width="67" height="30"><b><font color="#FF0000">��ַ˵����</font></b></td>
      <td width="208" height="30"> 
        <input type="text" name="txtname" size="33"            
          class="smallinput" maxlength="100" value="<%=rs("txtname")%>">
      </td>
    </tr>
    <tr> 
      <td width="72" align="right" height="8" bgcolor="#E1F4EE"><b><font color="#FF0000">�����С��</font></b></td>
      <td width="250" height="8" bgcolor="#E1F4EE"> 
        <!--<input type="text" name="system" size="60" class="smallinput" maxlength="100" value="Win98/2000+PWS4&NT+IIS4/5">-->
        <input type="text" name="size" size="33"             
          class="smallinput" maxlength="100" value="<%if rs("size")<>"" then%><%=rs("size")%><%else%>K<%end if%>">
      </td>
      <td width="67" height="8" bgcolor="#E1F4EE" align="right"> <b>������ʾ��</b> 
      </td>
      <td height="8" bgcolor="#E1F4EE" width="208">
        <input type="text" name="txtbb" size="33"            
          class="smallinput" maxlength="100" value="<%=rs("bb")%>">
        </td>
    </tr>
    <tr bgcolor="#E1F4EE"> 
      <td width="72" align="right" height="30" bgcolor="#E1F4EE"><b>�����Դ��</b></td>
      <td height="30" bgcolor="#E1F4EE" width="250"> 
        <input type="text" name="fromurl" size="33"             
          class="smallinput" maxlength="100" value="<%=rs("fromurl")%>">
      </td>
      <td height="30" bgcolor="#E1F4EE" width="67" align="right"><B>���ͼƬ��</B></td>
      <td height="30" bgcolor="#E1F4EE" width="208"> 
        <INPUT type="text" name="images" size="33"             
          class="smallinput" maxlength="160" value="<%=rs("images")%>">
      </td>
    </tr>
    <tr bgcolor="#E1F4EE"> 
      <td width="72" align="right" height="30"><b><font color="#FF0000">��Ȩ��ʽ��</font></b></td>
      <td height="30" colspan="3"> 
        <select name="order" size="1">
          <option  selected  value="������" name="orders">������</option>
                  </select>
        <!--<input type="text" name="order" size="60" class="smallinput" maxlength="100" value="���">-->
        ����: 
        <select name="xxyf" size="1">
          <option <%if rs("xxyf")="��������" then%> selected <%end if%> value="��������" name="xxyf">��������</option>
          <option <%if rs("xxyf")="��������" then%> selected <%end if%> value="��������" name="xxyf">��������</option>
          <option <%if rs("xxyf")="Ӣ��" then%> selected <%end if%> value="Ӣ��" name="xxyf">Ӣ��</option>
          <option <%if rs("xxyf")="����" then%> selected <%end if%> value="����" name="xxyf">����</option>
        </select>
        <!--<input type="text" name="xxyf" size="60" class="smallinput" maxlength="100" value="����">-->
        �Ƽ� 
        <input type="checkbox" name="hots" value="on" <%if int(rs("hots"))=1 then response.write "checked"%>>
        ��Ʒ 
        <input type="checkbox" name="hide" value="on" <%if int(rs("stop"))=1 then response.write "checked"%>>
        ��Ա 
        <select size="1" name="hy">
          <option value="0" selected name="hy">������</option>
        </select>
      </td>
    </tr>
    <tr bgcolor="#E1F4EE"> 
      <td width="72" align="right" valign="top" height="30"><b><font color="#FF0000">������</font>��</b></td>
      <td colspan="3" height="30"> 
        <textarea rows="6" name="txtnote" cols="82" class="smallarea"><%
	content=replace(rs("note"),"<br>",chr(13))
	content=replace(content,"&nbsp;"," ")
	response.write content
%></textarea>
      </td>
    </tr>
	<input type="hidden" name="txtfilename1" size="82"      
          class="smallinput" maxlength="100" value="<%=rs("filename1")%>">
		  <input type="hidden" name="txtname1" size="82"            
          class="smallinput" maxlength="100" value="<%=rs("txtname1")%>">
   <input type="hidden" name="txtfilename2" size="82"             
          class="smallinput" maxlength="100" value="<%=rs("filename2")%>">
 <input type="hidden" name="txtname2" size="82"            
          class="smallinput" maxlength="100" value="<%=rs("txtname2")%>">
    <input type="hidden" name="txtfilename3" size="82"             
          class="smallinput" maxlength="100" value="<%=rs("filename3")%>">
    <input type="hidden" name="txtname3" size="82"            
          class="smallinput" maxlength="100" value="<%=rs("txtname3")%>">
<input type="hidden" name="txtfilename4" size="82"             
          class="smallinput" maxlength="100" value="<%=rs("filename4")%>">
  <input type="hidden" name="txtname4" size="82"            
          class="smallinput" maxlength="100" value="<%=rs("txtname4")%>">

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