<html>

<!--#include file="articleconn.asp"-->
<%
dim sql
dim rs
dim typename
dim articleid

typename=request("typename")
articleid=request("id")
set rs=server.createobject("adodb.recordset")
sql="update learning set hits=hits+1 where articleID="&articleid
rs.open sql,conn,1,1   
sql="select * from learning where articleid="&request("id")
rs.open sql,conn,1,1
%>
<head>
<title><%=rs("title")%></title>
<link rel="stylesheet" href="css/article.css">
</head>

<body bgcolor="#FFFFFF">

<div align="center">
  <center>
    <table cellspacing="0" width="100%" border="1" bordercolorlight="#000000"
bordercolordark="#FFFFFF">
      <tr bgcolor="#CCCCCC"> 
        <td height="2" class="title" colspan="4" bgcolor="#B5D85E"> 
          <div align="center">��Ϣ��ѯϵͳ</div>
        </td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="20" class="title" colspan="4"> 
        
          <blockquote> 
            <div align="center"></div>
            <table width="65%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr> 
                <td colspan="2"> 
                  <div align="center">��ţ�<b><%=rs("articleid")%></b></div>
                </td>
              </tr>
              <tr> 
                <td colspan="2">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="2">��Ϣ���ƣ�<b><%=rs("title")%></b></td>
              </tr>
              <tr> 
                <td colspan="2">&nbsp;</td>
              </tr>
              <tr> 
                <td width="60%">��Ϣ���ͣ�<%=rs("type")%></td>
                <td width="40%">��Ϣ���ۣ�<font color="#FF0000"><%=rs("vote")%></font></td>
              </tr>
              <tr> 
                <td width="60%">&nbsp;</td>
                <td width="40%">&nbsp;</td>
              </tr>
              <tr> 
                <td width="60%">��Ϣ��С��<%=rs("big")%></td>
                <td width="40%">������ӣ�<a href="<%=rs("fromurl")%>"><%=rs("from")%></a></td>
              </tr>
              <tr> 
                <td width="60%">&nbsp;</td>
                <td width="40%">&nbsp;</td>
              </tr>
              <tr> 
                <td width="60%">ʱ�䣺<%=rs("dateandtime")%></td>
                <td width="40%">������<%=rs("hits")%></td>
              </tr>
              <tr> 
                <td width="60%">&nbsp;</td>
                <td width="40%">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="2">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="2"> 
                  <p>��Ϣ��飺</p>
                </td>
              </tr>
              <tr> 
                <td colspan="2"> 
                  <blockquote> 
                    <p><%=rs("content")%></p>
                  </blockquote>
                </td>
              </tr>
              <tr> 
                <td colspan="2"> 
                  <div align="center">��<a href="javascript:self.close()">�رձ�����</a>��</div>
                </td>
              </tr>
            </table>
            <blockquote>
              <p>&nbsp; </p>
            </blockquote>
          </blockquote>
        </td>
      </tr>
    </table>
    <div align="center"></div>
  </center>
</div>

</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
