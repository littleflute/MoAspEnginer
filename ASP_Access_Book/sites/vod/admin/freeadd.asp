<%
if session("admin")="" then
  response.redirect "admin.asp"
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
%><SCRIPT language = "JavaScript">
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
<SCRIPT language="javascript">
<!--
function CheckForm()
{
	document.myform.txtcontent.value=document.myform.doc_html.value;
	return true
}
//-->
</SCRIPT>
<body bgcolor=#009458>
<form method="POST" name="myform" action="adminsave.asp?action=add">
  <TABLE width="602" border="0" align="center" cellspacing="1">
    <TR align="center"> 
      <TD colspan="4" height="25" bgcolor="#145f74" width="594"><b><FONT color="#FFFFFF">�� 
        �� �� �� �� Ƶ</FONT></b></TD>
    </TR>
    <TR> 
      <TD align="right" width="76" nowrap bgcolor="#A5D0DC" height="30"><B><font color="#FF0000">ӰƬ����</font>��</B></TD>
      <TD width="225" bgcolor="#A5D0DC" height="30"> 
        <%
        sql = "select * from class"
        rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	response.write "���������Ŀ��"
	response.end
	else
%>
        <SELECT name="classid" onChange="changelocation(document.myform.classid.options[document.myform.classid.selectedIndex].value)" size="1">
          <OPTION selected value="">==��ѡ����==</OPTION>
          <% 
        do while not rs.eof
%>
          <OPTION value="<%=trim(rs("classid"))%>"><%=trim(rs("class"))%></OPTION>
          <%
        rs.movenext
        loop
	end if
        rs.close
        set rs = nothing
        conn.Close
        set conn = nothing
%>
        </SELECT>
      </TD>
      <TD align="right" width="76" bgcolor="#A5D0DC" height="30"><B>ѡ����ࣺ</B></TD>
      <TD width="225" bgcolor="#A5D0DC" height="30"> 
        <SELECT name="Nclassid">
          <OPTION selected value="">==��ѡ����==</OPTION>
        </SELECT>
      </TD>
    </TR>
    <tr> 
      <td width="76" align="right" height="30" nowrap bgcolor="#CCE6ED"><b><font color="#FF0000">ӰƬ����</font>��</b></td>
      <td width="225" height="30" bgcolor="#CCE6ED"> 
        <input type="text" name="txtshowname" size="33"              
          class="smallinput" maxlength="100">
      </td>
      <td width="76" height="30" bgcolor="#CCE6ED"> 
        <p align="right"><b>ӰƬ�汾��</b>
      </td>
      <td width="225" height="30" bgcolor="#CCE6ED"> 
        <input type="text" name="txtbb" size="33"             
          class="smallinput" maxlength="100">
      </td>
    </tr>
    <tr> 
      <td width="76" align="right" height="14" nowrap bgcolor="#A5D0DC">��Ӱ��ʽ:</td>
      <td width="450" height="14" colspan="3" bgcolor="#A5D0DC"> [<img border="0" src="../images/rm.gif"> 
        <input type="radio" value="rm" name="movie" checked>
        ]&nbsp; [<img border="0" src="../images/win.gif"> 
        <input type="radio" value="win" name="movie">
        ]&nbsp; [���ǵ�Ӱ 
        <input type="radio" value="" name="movie">
        ]&nbsp;&nbsp;&nbsp; ��������ߵ�Ӱ��ѡ�񲥷Ÿ�ʽ</td>
    </tr>
    <tr> 
      <td width="76" align="right" height="15" nowrap bgcolor="#A5D0DC">��Ա��</td>
      <td width="450" height="15" colspan="3" bgcolor="#A5D0DC">�Ƿ�Ʒ 
        <input type="checkbox" name="club" value="on">&nbsp;���ѡ���Ա��ֻ�л�Ա�ſɹۿ�</td>
    </tr>
    <tr> 
      <td width="76" align="right" height="30" nowrap bgcolor="#cce6ed"><b><font color="#FF0000">���ص�ַ1</font>��</b></td>
      <td width="450" height="30" colspan="3" bgcolor="#cce6ed"> 
        <input type="text" name="txtfilename" size="82"            
          class="smallinput" maxlength="100" value="<%=downdir1%>">
      </td>
    </tr>
    <tr> 
      <td width="76" align="right" height="30" nowrap bgcolor="#a5d0dc"><b>���ص�ַ2��</b></td>
      <td width="450" height="30" colspan="3" bgcolor="#a5d0dc"> 
        <input type="text" name="txtfilename1" size="82"               
          class="smallinput" maxlength="100" value="">
      </td>
    </tr>
    <tr bgcolor="#cce6ed"> 
      <td width="76" align="right" height="30" nowrap><b>���ص�ַ3��</b></td>
      <td width="450" height="30" colspan="3"> 
        <input type="text" name="txtfilename2" size="82"               
          class="smallinput" maxlength="100" value="">
      </td>
    </tr>
    <tr bgcolor="#cce6ed"> 
      <td width="76" align="right" height="30" nowrap bgcolor="#A5D0DC"><b>�Ƽ��ȣ�</b></td>
      <td width="450" height="30" colspan="3" bgcolor="#A5D0DC"> 1��. 
        <input type="radio" value="1" name="hot">
        &nbsp;&nbsp;&nbsp; 2��. 
        <input type="radio" value="2" name="hot">
        &nbsp;&nbsp;&nbsp; 3��. 
        <input type="radio" value="3" name="hot" checked>
        &nbsp;&nbsp;&nbsp; 4��. 
        <input type="radio" value="4" name="hot">
        &nbsp;&nbsp;&nbsp; 5��. 
        <input type="radio" value="5" name="hot">
        <!--<input type="text" name="hot" size="60" class="smallinput" maxlength="100" value="3">-->
      </td>
    </tr>
    <tr> 
      <td width="76" align="right" height="30" nowrap bgcolor="#CCE6ED"><b>ӰƬ��С��</b></td>
      <td width="225" height="30" bgcolor="#CCE6ED"> 
        <!--<input type="text" name="system" size="60" class="smallinput" maxlength="100" value="Win98/2000+PWS4&NT+IIS4/5">-->
        <input type="text" name="size" size="20"               
          class="smallinput" maxlength="100" value="K">
      </td>
      <td width="76" height="30" bgcolor="#CCE6ED" align="right"> 
        <p align="right"><b>���л�����</b> 
      </td>
      <td width="225" height="30" bgcolor="#CCE6ED"> 
        <select name="system" size="1">
          <option value="Win9x/WinNT/Win2000/WinME" name="system">Win9x/WinNT/Win2000/WinME</option>
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
      <td width="76" align="right" height="30" nowrap bgcolor="#A5D0DC"><b>������ӣ�</b></td>
      <td width="225" height="30" bgcolor="#A5D0DC"> 
        <input type="text" name="fromurl" size="33"               
          class="smallinput" maxlength="100" value="">
      </td>
      <td width="76" height="30" align="right" bgcolor="#A5D0DC"><B>ӰƬͼƬ��</B></td>
      <td width="225" height="30" bgcolor="#A5D0DC"> 
        <INPUT type="text" name="images" size="33"               
          class="smallinput" maxlength="160">
      </td>
    </tr>
    <tr bgcolor="#cce6ed"> 
      <td width="76" align="right" height="30" nowrap><b>��Ȩ��ʽ��</b></td>
      <td width="450" height="30" colspan="3"> 
        <select name="order" size="1">
          <option selected value="����ӰƬ" name="order">����ӰƬ</option>
          <option value="���ӰƬ" name="order">���ӰƬ</option>
          <option value="��ҵӰƬ" name="order">��ҵӰƬ</option>
        </select>
        <!--<input type="text" name="order" size="60" class="smallinput" maxlength="100" value="���">-->
        �Ƿ��Ƽ� 
        <input type="checkbox" name="hots" value="on">
        �Ƿ�Ʒ 
        <input type="checkbox" name="hide" value="on">
        <font color="#FF0000">�����ѡ����ü���ͼƬ��ַ������ҳ��ʾ��</font></td>
    </tr>
    <tr> 
      <td width="76" align="right" valign="top" nowrap bgcolor="#a5d0dc"><b><font color="#FF0000">ӰƬ��飺</font></b></td>
      <td width="450" colspan="3" bgcolor="#a5d0dc"> <FONT color="#FF0000"> 
        <textarea rows="6" name="txtnote" cols="82" class="smallarea"></textarea>
        </FONT></td>
    </tr>
  </TABLE>               
<div align="center"><center><p><input type="submit" value=" �� �� "               
  name="cmdok" class="buttonface">&nbsp; <input type="reset" value=" �� �� "               
  name="cmdcancel" class="buttonface"></p>               
  </center></div>               
</form>               
</body>               
</html>               
