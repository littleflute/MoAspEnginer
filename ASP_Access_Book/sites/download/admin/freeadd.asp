<!--#include file="conn.asp"-->
<!--#include file="../inc/const.asp"--><%
if session("admin")="" then
  response.redirect "admin.asp"
end if
%> 
<HTML><HEAD>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">


<TITLE>添 加 下 载 程 序</TITLE>
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
        <input type="submit" value=" 添 加 " name="cmdok2">
        &nbsp; 
        <input type="reset" value=" 清 除 " name="cmdcancel2">
      </TD>
    </TR>
    <TR> 
      <TD align="right" width="72" nowrap height="30" bgcolor="E1F4EE"><B><FONT color="#FF0000">软件类型</FONT></B></TD>
      <TD width="231" height="30" bgcolor="E1F4EE"> 
        <%
        sql = "select * from class"
        rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	response.write "请先添加栏目。"
	response.end
	else
%>
        <SELECT name="classid" onChange="changelocation(document.myform.classid.options[document.myform.classid.selectedIndex].value)" size="1">
          <OPTION selected value>==请选类型==</OPTION>
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
      <TD align="right" width="70" height="30" bgcolor="E1F4EE"><B><font color="#FF0000">选择分类</font></B></TD>
      <TD width="208" height="30" bgcolor="E1F4EE"> 
        <SELECT name="Nclassid">
          <OPTION selected value>==请选分类==</OPTION>
        </SELECT>
      </TD>
    </TR>
    <TR> 
      <TD width="72" align="right" height="30" nowrap bgcolor="E1F4EE"><B><FONT color="#FF0000">软件名称</FONT></B></TD>
      <TD width="231" height="30" bgcolor="E1F4EE"> 
        <INPUT type="text" name="txtshowname" size="30" class="smallinput" maxlength="100">
      </TD>
      <TD width="70" height="30" bgcolor="E1F4EE" align="right"> 
        <P align="right"><B><font color="#FF0000">运行环境</font></B> 
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
      <TD width="72" align="right" height="30" nowrap bgcolor="E1F4EE"><font color="#FF0000">格式:</font></TD>
      <TD width="231" height="30" bgcolor="E1F4EE">[软件 
        <input type="radio" value name="movie" checked>
        ]</TD>
      <TD width="70" height="30" bgcolor="E1F4EE" align="right"> <b><font color="#FF0000">推荐度</font></b></TD>
      <TD width="208" height="30" bgcolor="E1F4EE"> 
        <select name="hot">
          <option value="1">1星级</option>
          <option value="2">2星级</option>
          <option value="3">3星级</option>
          <option value="4" selected>4星级</option>
          <option value="5">5星级</option>
        </select>
      </TD>
    </TR>
    <TR> 
      <TD width="72" align="right" height="30" nowrap bgcolor="E1F4EE"><B><FONT color="#FF0000">下载地址1</FONT><a href="../admin_uploadfile.asp"><br>
        </a></B></TD>
      <TD width="231" height="30" bgcolor="E1F4EE"> 
        <INPUT type="text" name="txtfilename" size="30" class="smallinput" maxlength="100" value="<%=downdir1%>">
      </TD>
      <TD width="70" height="30" bgcolor="E1F4EE" align="right"><b><font color="#FF0000">地址说明</font></b></TD>
      <TD width="208" height="30" bgcolor="E1F4EE"> 
        <input type="text" name="txtname" size="30" class="smallinput" maxlength="100" value="<%=downdir1%>点击下载">
      </TD>
    </TR>
    <TR> 
      <TD width="72" align="right" height="30" nowrap bgcolor="E1F4EE"><B><font color="#FF0000">软件大小</font></B></TD>
      <TD width="231" height="30" bgcolor="E1F4EE"> 
        <INPUT type="text" name="size" size="20" class="smallinput" maxlength="100" value="不详">
      </TD>
      <TD width="70" height="30" align="right" bgcolor="E1F4EE"> 
        <P><B>程序演示</B> 
      </TD>
      <TD width="208" height="30" bgcolor="E1F4EE"> 
        <input type="text" name="txtbb" size="30" class="smallinput" maxlength="100">
      </TD>
    </TR>
    <TR> 
      <TD width="72" align="right" height="30" nowrap bgcolor="E1F4EE"><B>相关链接</B></TD>
      <TD width="231" height="30" bgcolor="E1F4EE"> 
        <INPUT type="text" name="fromurl" size="30" class="smallinput" maxlength="100">
      </TD>
      <TD width="70" height="30" align="right" bgcolor="E1F4EE"><B>软件图片</B></TD>
      <TD width="208" height="30" bgcolor="E1F4EE"> 
        <INPUT type="text" name="images" size="25" class="smallinput" maxlength="160">
      </TD>
    </TR>
    <TR> 
      <TD width="72" align="right" height="30" nowrap bgcolor="E1F4EE"><B><font color="#FF0000">授权形式</font></B></TD>
      <TD height="30" colspan="3" bgcolor="E1F4EE"> 
        <select name="order" size="1">
          <option value="免费软件" selected>免费软件</option>
        </select>
        语言 
        <SELECT name="xxyf" size="1">
          <OPTION value="简体中文" name="order" selected>简体中文</OPTION>
          <OPTION value="繁体中文">繁体中文</OPTION>
          <OPTION value="英文" name="order">英文</OPTION>
          <OPTION value="其他" name="order">其他</OPTION>
        </SELECT>
        推荐 
        <input type="checkbox" name="hots" value="on">
        精品 
        <input type="checkbox" name="hide" value="on">
        级别 
        <select size="1" name="hy">
          <option  value="0">免费软件</option>
        </select>
      </TD>
    </TR>
    <TR> 
      <TD width="72" align="right" valign="top" nowrap bgcolor="E1F4EE"><B><FONT color="#FF0000">软件简介</FONT></B></TD>
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
             <P><INPUT type="submit" value=" 添 加 " name="cmdok" class="buttonface">&nbsp; <INPUT type="reset" value=" 清 除 " name="cmdcancel" class="buttonface"></P>
             </CENTER></DIV>
</FORM>
</BODY></HTML>