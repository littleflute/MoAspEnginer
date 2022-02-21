<!--#include file="conn.asp"-->
<style type="text/css">
<!--
a:visited{text-transform: none; text-decoration: none; font-style: normal; font-weight: normal; font-variant: normal; color: #000000}
a:hover{text-decoration:underline; color: #FF3333}
a:link{text-transform: none; text-decoration: none; font-weight: normal; font-variant: normal; color: #000000}
.p9 { font-size: 9pt; line-height: 13pt; text-decoration: none}
-->
</style>
<% If (not session("AdminUID")="") and session("Adminflag")="0" Then %>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid; background-color: #e3f1d1;">
    <tr class="p9">
    <td width="469" height="16"><a href="../index.asp">论坛论坛</a> →<a href="oneedit.asp">超级管理</a> 
      →→ <a href="sql.asp">执行sql语句</a> </td>
    <td width="289">管理菜单：| <a href="board.asp?boardid=0">查看公告</a> <a href="oneedit.asp?pub=yes">| 
      管理公告 </a> | <a href="announce.asp?boardid=0">发布公告 </a> |</td>
  </tr>
</table>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="64%">&nbsp;</td>
    <td width="36%">&nbsp;</td>
  </tr>
  <form action=sql.asp method="post">
    <tr> 
      <td> <span class="p9"><font color="#FF0000">待执行SQL语句(警告：执行语句要万分小心!) ××<a href="date/datebackup.asp">请务必先使用数据库备份</a>××<br>
        使用delete语句时务必加好条件，不用使用如：“delete from 表名”<br>
        </font></span> 
        <textarea name=GBL_EXEString rows=2 cols=61 class=fmtxtra><%If GBL_EXEString <> "" Then Response.Write VbCrLf & htmlEncode(GBL_EXEString)%></textarea> 
        <input name=submitflag type=hidden value="Dieos9xsl29LO_8"> <input name="submit" type=submit class=p9 value="执行"> 
        <input name="reset" type=reset class=p9 value="取消"> </td>
      <td class="p9">常用sql语句<br>
        【更新】update 表名 set 字段1=值,字段2=值<br>
        例:update uers set flag=null where id='aa'<br> <br>
        【删除】delete from 表名 where 条件<br>
        例:delete from news where bbs_id=&quot;&amp;tid</td>
    </tr>
  </form>
  <tr> 
    <td colspan="2" class="p9"> 
      <%  
	   GBL_EXEString = Request("GBL_EXEString")
	   If GBL_EXEString <> "" Then
			On Error Resume Next
			Dim RowCount,Rs
			Dim Time1,Time2
			Time1=Timer
			Dim GBL_EXEString
			conn.ExeCute(GBL_EXEString)
			Time2=Timer
			
	'		If DEF_UsedDataBase = 0 Then
    '			Set Rs = conn.Execute("select @@rowcount")
	' 			RowCount = Rs(0)
	' 			Rs.Close
	'		Else
	'			RowCount = "<font color=ff0000>未知</font>"
	'		End if
	
			Set Rs = Nothing
			if err.number<>0 then
				Response.Write "<br><span style=""FONT-FAMILY: 宋体; FONT-SIZE: 12px;""><font color=ff0000><b>数据库命令操作失败：</b></font><p>"&err.description & "</span>"
			Else
				Response.Write "<br><span style=""FONT-FAMILY: 宋体; FONT-SIZE: 12px;""><font color=008800><b>下列数据库命令操作成功，共影响<font color=ff0000>" & RowCount & "</font>行数据，耗时" & (Time2-Time1)*1000 & "毫秒!</b></font></span><hr size=1>" & "<hr size=1>" & VbCrLf
			End If		
		Else
			Response.Write "<br><font color=ff0000><b>命令不能为空!</b></font>"
		End If
%><br><font color="#FF0000">【sql语句参考】 </font><br>
      【将一个帖子换一个版面】update news set boardid=4 where bbs_id=418 or parent_id=418 
      <br>
      【设置某用户为某版斑竹】[首页] update board set banzhu='风影' where boardid=2 [设置权限]update 
      uers set flag=2 where id='风影' 
      <hr size="1" noshade></td></tr>
  <tr> 
    <td colspan="2">&nbsp;</td>
  </tr>
</table>
<% Else %>

<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="center">&nbsp;你没有管理权限或请重新登陆！<a href="javascript:history.go(-1)">请点击这里返回！</a> 
	</td>
  </tr>
</table>
<% End If %>