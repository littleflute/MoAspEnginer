<% 		
		set rs_menu=Server.CreateObject("ADODB.Recordset")
		sql_menu="select count(*) as menu_user from uers"
		rs_menu.Open sql_menu,db,1,3 
		menu_user= rs_menu("menu_user")
		rs_menu.close
		sql_menu="select count(*) as menu_tit from news where layer=1"
		rs_menu.Open sql_menu,db,1,1
		menu_tit=rs_menu("menu_tit")
		rs_menu.close
		sql_menu="select count(*) as menu_title from news"
		rs_menu.Open sql_menu,db,1,1
		menu_title=rs_menu("menu_title")
		rs_menu.close
		sql_menu="select count(*) as menu_today from news where year(submit_date)="&year(now())&"  and month(submit_date)="&month(now())&" and day(submit_date)="&day(now())
		rs_menu.Open sql_menu,db,1,1
		menu_today=rs_menu("menu_today")
		rs_menu.close
		set rs_menu=nothing
%>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid; background-color: #e3f1d1;">
  <tr> 
    <td width="473">��̳���� <%= menu_user %> λע���Ա , ����������<%= menu_tit %> , ����������<%= menu_title %>,����������<font color="#FF0000"><%= menu_today %></font></td>
    <td width="285">| <a href="list.asp?order=1">��̳����</a> | <a href="list.asp?order=2">���Ż���</a> 
      | <a href="uerlist.asp?id=2">��������</a> | <a href="uerlist.asp">�û��б�</a> |</td>
  </tr>
</table>
