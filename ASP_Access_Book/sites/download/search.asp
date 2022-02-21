<HTML>
<HEAD>
<!--#include file="conn.asp"-->

<META http-equiv="Content-Language" content="zh-cn">
<TITLE>搜索引擎</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
</HEAD>
<!--#include file="head.asp"-->
<DIV align="center">
  <DIV align="center"> 
    <!--#include file="lanmu.asp"-->
    <br>
    <TABLE  height="66" cellSpacing="0" cellPadding="0" width="770" bgColor="#ffffff" border="0">
           
          <TR> 
            <TD vAlign="top" align="middle" width="100%" height="154"><BR>
              <CENTER>
                <P></P>
              </CENTER>
              <TABLE style="BORDER-RIGHT: 1px dotted; BORDER-TOP: 1px dotted; BORDER-LEFT: 1px dotted; BORDER-BOTTOM: 1px dotted; BORDER-COLLAPSE: collapse" height="11" cellSpacing="0" cellPadding="3" width="400" bgColor="#FFCC00" border="0" align="center">
                 
                <TR> 
                  <TD width="80" bgColor="#ffffff" height="10"><IMG src="images/search1.gif" border="0" width="79" height="32"></TD>
                  <TD width="306" bgColor="#ffffff" height="10"> 
                    <P align="center"><b>搜索引擎</b> 
                  </TD>
                </TR>
                <TR> 
                  <FORM method="post" name="myform" action="query.asp">
                    <TD width="394" colSpan="2" height="77">软件搜索： 
                      <SELECT size="1" name="action" style="color:#008080;font-size: 9pt">
                        <OPTION value="title">程序名称</OPTION>
                        <OPTION value="content">程序简介</OPTION>
                      </SELECT>
                      <SELECT name="classid" size="1" style="color:#008080;font-size: 9pt">
                       <%
  dim rs
  set rs = server.createobject("adodb.recordset")
 sql = "select * from class"                        
        rs.open sql,conn,1,1                        
        do while not rs.eof                        
        %>
                        <OPTION value="<%=trim(rs("classid"))%>"><%=trim(rs("class"))%></OPTION>
                        <%                        
        rs.movenext                        
        loop                        
        rs.close                        
        %>
                      </SELECT>
                      <INPUT type="text" name="keyword" class="smallinput" size="14" value="关键字" maxlength="50" style="color:#008080;font-size: 9pt">
                      <INPUT type="submit" name="Submit2" value="搜索" style="height='21'">
                    </TD>
                  </FORM>
                </TR>
                 
              </TABLE>
            </TD>
          </TR>
           
        </TABLE>
        
    <br>
  </DIV>
</DIV>
<!--#include file="foot.asp"-->
</BODY>
</HTML>