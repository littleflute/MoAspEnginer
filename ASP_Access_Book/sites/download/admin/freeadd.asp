<!--#include file="conn.asp"-->
<!--#include file="../inc/const.asp"--><%
if session("admin")="" then
  response.redirect "admin.asp"
end if
%> 
<HTML><HEAD>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">


<TITLE>�� �� �� �� �� ��</TITLE>
<LINK rel="stylesheet" type="text/css" href="style.css">
</HEAD>
<%
dim rs
dim sql
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from Nclass order by Nclassid asc"
rs.open sql,conn,1,1
%>
<SCRIPT language="JavaScript">
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
</SCRIPT> <SCRIPT language="javascript">
<!--
function CheckForm()
{
	document.myform.txtcontent.value=document.myform.doc_html.value;
	return true
}
//-->
</SCRIPT>
<BODY leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF">
<FORM method="POST" name="myform" action="adminsave.asp?action=add">
  <TABLE width="578" border="1" align="center" cellspacing="0" bordercolor="#39867B"  bordercolordark=#FFFFFF cellpadding="0">
    <TR align="center" bgcolor="E1F4EE"> 
      <TD colspan="4" height="20"> 
        <input type="submit" value=" �� �� " name="cmdok2">
        &nbsp; 
        <input type="reset" value=" �� �� " name="cmdcancel2">
      </TD>
    </TR>
    <TR> 
      <TD align="right" width="72" nowrap height="30" bgcolor="E1F4EE"><B><FONT color="#FF0000">�������</FONT></B></TD>
      <TD width="231" height="30" bgcolor="E1F4EE"> 
        <%
        sql = "select * from class"
        rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	response.write "���������Ŀ��"
	response.end
	else
%>
        <SELECT name="classid" onChange="changelocation(document.myform.classid.options[document.myform.classid.selectedIndex].value)" size="1">
          <OPTION selected value>==��ѡ����==</OPTION>
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
      <TD align="right" width="70" height="30" bgcolor="E1F4EE"><B><font color="#FF0000">ѡ�����</font></B></TD>
      <TD width="208" height="30" bgcolor="E1F4EE"> 
        <SELECT name="Nclassid">
          <OPTION selected value>==��ѡ����==</OPTION>
        </SELECT>
      </TD>
    </TR>
    <TR> 
      <TD width="72" align="right" height="30" nowrap bgcolor="E1F4EE"><B><FONT color="#FF0000">�������</FONT></B></TD>
      <TD width="231" height="30" bgcolor="E1F4EE"> 
        <INPUT type="text" name="txtshowname" size="30" class="smallinput" maxlength="100">
      </TD>
      <TD width="70" height="30" bgcolor="E1F4EE" align="right"> 
        <P align="right"><B><font color="#FF0000">���л���</font></B> 
      </TD>
      <TD width="208" height="30" bgcolor="E1F4EE"> 
        <select name="system" size="1">
          <option value="Win9x/WinNT/Win2000/WinXP/WinME" selected>Win9x/WinNT/Win2000/WinXP/WinME</option>
          <option value="Win9x/WinNT/Win2000/WinME">Win9x/WinNT/Win2000/WinME</option>
          <option value="Win9x/WinNT/Win2000">Win9x/WinNT/Win2000</option>
          <option value="Win9x/WinNT">Win9x/WinNT</option>
          <option value="Win9x/WinME">Win9x/WinME</option>
        </select>
      </TD>
    </TR>
    <TR> 
      <TD width="72" align="right" height="30" nowrap bgcolor="E1F4EE"><font color="#FF0000">��ʽ:</font></TD>
      <TD width="231" height="30" bgcolor="E1F4EE">[��� 
        <input type="radio" value name="movie" checked>
        ]</TD>
      <TD width="70" height="30" bgcolor="E1F4EE" align="right"> <b><font color="#FF0000">�Ƽ���</font></b></TD>
      <TD width="208" height="30" bgcolor="E1F4EE"> 
        <select name="hot">
          <option value="1">1�Ǽ�</option>
          <option value="2">2�Ǽ�</option>
          <option value="3">3�Ǽ�</option>
          <option value="4" selected>4�Ǽ�</option>
          <option value="5">5�Ǽ�</option>
        </select>
      </TD>
    </TR>
    <TR> 
      <TD width="72" align="right" height="30" nowrap bgcolor="E1F4EE"><B><FONT color="#FF0000">���ص�ַ1</FONT><a href="../admin_uploadfile.asp"><br>
        </a></B></TD>
      <TD width="231" height="30" bgcolor="E1F4EE"> 
        <INPUT type="text" name="txtfilename" size="30" class="smallinput" maxlength="100" value="<%=downdir1%>">
      </TD>
      <TD width="70" height="30" bgcolor="E1F4EE" align="right"><b><font color="#FF0000">��ַ˵��</font></b></TD>
      <TD width="208" height="30" bgcolor="E1F4EE"> 
        <input type="text" name="txtname" size="30" class="smallinput" maxlength="100" value="<%=downdir1%>�������">
      </TD>
    </TR>
    <TR> 
      <TD width="72" align="right" height="30" nowrap bgcolor="E1F4EE"><B><font color="#FF0000">�����С</font></B></TD>
      <TD width="231" height="30" bgcolor="E1F4EE"> 
        <INPUT type="text" name="size" size="20" class="smallinput" maxlength="100" value="����">
      </TD>
      <TD width="70" height="30" align="right" bgcolor="E1F4EE"> 
        <P><B>������ʾ</B> 
      </TD>
      <TD width="208" height="30" bgcolor="E1F4EE"> 
        <input type="text" name="txtbb" size="30" class="smallinput" maxlength="100">
      </TD>
    </TR>
    <TR> 
      <TD width="72" align="right" height="30" nowrap bgcolor="E1F4EE"><B>�������</B></TD>
      <TD width="231" height="30" bgcolor="E1F4EE"> 
        <INPUT type="text" name="fromurl" size="30" class="smallinput" maxlength="100">
      </TD>
      <TD width="70" height="30" align="right" bgcolor="E1F4EE"><B>���ͼƬ</B></TD>
      <TD width="208" height="30" bgcolor="E1F4EE"> 
        <INPUT type="text" name="images" size="25" class="smallinput" maxlength="160">
      </TD>
    </TR>
    <TR> 
      <TD width="72" align="right" height="30" nowrap bgcolor="E1F4EE"><B><font color="#FF0000">��Ȩ��ʽ</font></B></TD>
      <TD height="30" colspan="3" bgcolor="E1F4EE"> 
        <select name="order" size="1">
          <option value="������" selected>������</option>
        </select>
        ���� 
        <SELECT name="xxyf" size="1">
          <OPTION value="��������" name="order" selected>��������</OPTION>
          <OPTION value="��������">��������</OPTION>
          <OPTION value="Ӣ��" name="order">Ӣ��</OPTION>
          <OPTION value="����" name="order">����</OPTION>
        </SELECT>
        �Ƽ� 
        <input type="checkbox" name="hots" value="on">
        ��Ʒ 
        <input type="checkbox" name="hide" value="on">
        ���� 
        <select size="1" name="hy">
          <option  value="0">������</option>
        </select>
      </TD>
    </TR>
    <TR> 
      <TD width="72" align="right" valign="top" nowrap bgcolor="E1F4EE"><B><FONT color="#FF0000">������</FONT></B></TD>
      <TD colspan="3" bgcolor="E1F4EE"><FONT color="#FF0000"> 
        <TEXTAREA rows="6" name="txtnote" cols="75" class="smallarea"></TEXTAREA>
        </FONT></TD>
    </TR>
    <input type="hidden" name="txtfilename1" size="82" class="smallinput" maxlength="100" value>
    <input type="hidden" name="txtname1" size="82" class="smallinput" maxlength="100" value="<%=downdir1%>">
    <input type="hidden" name="txtfilename2" size="82" class="smallinput" maxlength="100" value>
    <input type="hidden" name="txtname2" size="82" class="smallinput" maxlength="100" value>
    <input type="hidden" name="txtfilename3" size="82" class="smallinput" maxlength="100" value>
    <input type="hidden" name="txtname3" size="82" class="smallinput" maxlength="100" value>
    <input type="hidden" name="txtfilename4" size="82" class="smallinput" maxlength="100" value>
    <input type="hidden" name="txtname4" size="82" class="smallinput" maxlength="100" value>
  </TABLE>
  <DIV align="center">
             <CENTER>
             <P><INPUT type="submit" value=" �� �� " name="cmdok" class="buttonface">&nbsp; <INPUT type="reset" value=" �� �� " name="cmdcancel" class="buttonface"></P>
             </CENTER></DIV>
</FORM>
</BODY></HTML>