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
    <td width="469" height="16"><a href="../index.asp">��̳��̳</a> ��<a href="oneedit.asp">��������</a> 
      ���� <a href="sql.asp">ִ��sql���</a> </td>
    <td width="289">����˵���| <a href="board.asp?boardid=0">�鿴����</a> <a href="oneedit.asp?pub=yes">| 
      ������ </a> | <a href="announce.asp?boardid=0">�������� </a> |</td>
  </tr>
</table>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="64%">&nbsp;</td>
    <td width="36%">&nbsp;</td>
  </tr>
  <form action=sql.asp method="post">
    <tr> 
      <td> <span class="p9"><font color="#FF0000">��ִ��SQL���(���棺ִ�����Ҫ���С��!) ����<a href="date/datebackup.asp">�������ʹ�����ݿⱸ��</a>����<br>
        ʹ��delete���ʱ��ؼӺ�����������ʹ���磺��delete from ������<br>
        </font></span> 
        <textarea name=GBL_EXEString rows=2 cols=61 class=fmtxtra><%If GBL_EXEString <> "" Then Response.Write VbCrLf & htmlEncode(GBL_EXEString)%></textarea> 
        <input name=submitflag type=hidden value="Dieos9xsl29LO_8"> <input name="submit" type=submit class=p9 value="ִ��"> 
        <input name="reset" type=reset class=p9 value="ȡ��"> </td>
      <td class="p9">����sql���<br>
        �����¡�update ���� set �ֶ�1=ֵ,�ֶ�2=ֵ<br>
        ��:update uers set flag=null where id='aa'<br> <br>
        ��ɾ����delete from ���� where ����<br>
        ��:delete from news where bbs_id=&quot;&amp;tid</td>
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
	'			RowCount = "<font color=ff0000>δ֪</font>"
	'		End if
	
			Set Rs = Nothing
			if err.number<>0 then
				Response.Write "<br><span style=""FONT-FAMILY: ����; FONT-SIZE: 12px;""><font color=ff0000><b>���ݿ��������ʧ�ܣ�</b></font><p>"&err.description & "</span>"
			Else
				Response.Write "<br><span style=""FONT-FAMILY: ����; FONT-SIZE: 12px;""><font color=008800><b>�������ݿ���������ɹ�����Ӱ��<font color=ff0000>" & RowCount & "</font>�����ݣ���ʱ" & (Time2-Time1)*1000 & "����!</b></font></span><hr size=1>" & "<hr size=1>" & VbCrLf
			End If		
		Else
			Response.Write "<br><font color=ff0000><b>�����Ϊ��!</b></font>"
		End If
%><br><font color="#FF0000">��sql���ο��� </font><br>
      ����һ�����ӻ�һ�����桿update news set boardid=4 where bbs_id=418 or parent_id=418 
      <br>
      ������ĳ�û�Ϊĳ�����[��ҳ] update board set banzhu='��Ӱ' where boardid=2 [����Ȩ��]update 
      uers set flag=2 where id='��Ӱ' 
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
    <td align="center">&nbsp;��û�й���Ȩ�޻������µ�½��<a href="javascript:history.go(-1)">�������ﷵ�أ�</a> 
	</td>
  </tr>
</table>
<% End If %>