<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3"> <% 
if  Request.QueryString("boardid")="" then
response.Redirect("index.asp")
else
boardid=Trim(Request.QueryString("boardid"))
end if
sql= "select * from board where boardid="&boardid
 set rs_board=db.execute(sql)
 banzhu =rs_board("banzhu")
%> <table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid; background-color: #e3f1d1;">
        <tr> 
          <td height="16"><a href="index.asp">论坛</a> → <%= rs_board("parent") %>→ <a href="board.asp?boardid=<%=boardid%>"><%= rs_board("name") %></a></td>
        </tr>
      </table>
      <% 
rs_board.close
set rs_board=nothing
 %> </td>
  </tr>
  <tr> 
    <td width="175">&nbsp;</td>
    <td width="425">&nbsp;</td>
    <td width="175"><iframe src=duanalert.asp  height=18 width=140 align="right"  Frameborder=No Border=0 Marginwidth=0 Marginheight=0 Scrolling=No></iframe></td>
  </tr>
  <tr> 
    <td colspan="3"> 
<%  
sqlgg="select * from news where boardid=0 and layer=1 order by submit_date desc"
set rsgg = db.execute(sqlgg)
%>	 
<SCRIPT LANGUAGE='JavaScript' SRC='js/fader.js' TYPE='text/javascript'></script>
<SCRIPT LANGUAGE='JavaScript' TYPE='text/javascript'>
prefix="";
arNews = ["<% If not (rsgg.eof and rsgg.bof)Then %><a href='particular.asp?bbs_id=<%= rsgg("bbs_id") %>&boardid=0' target=_blank><b><%= rsgg("title") %></b></a> [<%= rsgg("submit_date") %>]<% Else %>当前还没有公告<% End If %>",""," <% If not rsgg.eof Then
 rsgg.movenext
If not rsgg.eof Then %><a href='particular.asp?bbs_id=<%= rsgg("bbs_id") %>&boardid=0' target=_blank><b><%= rsgg("title") %></b></a> [<%= rsgg("submit_date") %>]<% Else %>论坛欢迎你的光临<% End If %> <% Else %>论坛欢迎你的光临<% End If %>","",
" <% If not rsgg.eof Then
 rsgg.movenext
If not rsgg.eof Then %><a href='particular.asp?bbs_id=<%= rsgg("bbs_id") %>&boardid=0' target=_blank><b><%= rsgg("title") %></b></a> [<%= rsgg("submit_date") %>]<% Else %>严禁恶意使用粗言秽语，违者经劝告无效，立即封ID<% End If %>
<% Else %>严禁恶意使用粗言秽语，违者经劝告无效，立即封ID<% End If %>",""]
</SCRIPT>
<div id="elFader" style="position:relative;visibility:hidden; height:16" ></div>
<% 
rsgg.close
set rsgg=nothing
 %>	
	</td>
  </tr>
</table>
