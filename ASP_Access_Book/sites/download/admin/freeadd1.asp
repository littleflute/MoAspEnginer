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
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>添 加 下 载 程 序</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<%
dim rs
dim sql
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from news "
rs.open sql,conn,1,1
%><SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
        count = 0
        do while not rs.eof 
        %>

        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.newsid.length = 0; 

    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.newsid.options[document.myform.newssid.length] = new Option(subcat[i][0], subcat[i][2]);
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
<body bgcolor=#E1F4EE>
<form method="POST" name="myform" action="adminsave1.asp?action=add">
  <TABLE width="80%" border="1" align="center" cellspacing="0" bordercolor="#39867B" cellpadding="0" bordercolordark=#FFFFFF>
    <TR align="center"> 
      <TD colspan="4" height="25" bgcolor="#E1F4EE"><b><FONT color="#000000">添 
        加 站 内 新 闻</FONT></b></TD>
    </TR>    
    <tr> 
      <td width="15%" align="right" valign="top" nowrap bgcolor="#E1F4EE"><b>新闻内容：</b></td>
      <td width="85%" colspan="3" bgcolor="#E1F4EE"><font color="#FF0000"> 
        <textarea rows="6" name="newsnote" cols="60" class="smallarea"></textarea>
        </font></td>
    </tr>
  </TABLE>
<div align="center"><center><p><input type="submit" value=" 添 加 "
  name="cmdok" class="buttonface">&nbsp; <input type="reset" value=" 清 除 "
  name="cmdcancel" class="buttonface"></p>
  </center></div>
</form>
</body>
</html>
