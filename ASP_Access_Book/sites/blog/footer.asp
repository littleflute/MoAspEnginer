            
<table cellspacing=0 cellpadding=0 align="center" width=776 border=0 background="images/blog_main.gif" class="wordbreak">
  <tr> <td width="8"> </td>
                <td width="700" bgcolor="#93936F">

<table width="793" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="749" height="59" align="center" class="td20"> 
                  <table width="740" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="16" align="right">&nbsp;</td>
                      <td width="724" bgcolor="#93936F"><div align="right"><font face="Verdana, Arial, Helvetica, sans-serif">Copyright ? 2006&nbsp;<%=siteName%>&nbsp;  <span style="font-size:9px;">
                      
                      </span> 
                      </font></div></td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
			
				</td>
               <td width="8"> </td></tr>
            </table>

<table width="776" height="5" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#FFFFFF"></td>
  </tr>
</table>
</body>
</html><%
IF TypeName(Conn)<>"Nothing" Then
	Conn.Close
	Set Conn=Nothing
Session.CodePage=936
End IF%>