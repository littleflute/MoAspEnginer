<HTML><HEAD><TITLE><%=toptitle%></TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="xsmz.css" rel=stylesheet>
<META content="MSHTML 5.00.3532.300" name=GENERATOR>
<meta name="KEYWords" contect="������,����">
<meta name="DEscription" contect="����,������,������">
<meta name="Author" contect="19452949">
<script LANGUAGE="JavaScript">
<!--
// Scrolling message settings
var MessageText = "��ӭ����--��������վ"
var DisplayLength = 260
var pos = 1 - DisplayLength;
function ScrollInStatusBar(){
var scroll = "";
pos++;
if (pos == MessageText.length) pos = 1 - DisplayLength;
if (pos<0) {
for (var i=1; i<=Math.abs(pos); i++)
scroll = scroll + "";
scroll = scroll + MessageText.substring(0, DisplayLength - i + 1);
}
else
scroll = scroll + MessageText.substring(pos, pos + DisplayLength);
window.status = scroll;
//Scrolling speed
setTimeout("ScrollInStatusBar()",50);
}
ScrollInStatusBar()
//-->
</script>
</HEAD>
<BODY background=images/bg.gif leftMargin=0 topMargin=0>
<SCRIPT>
function checkForm(frm)
{
	if (frm.keyword.value == null || frm.keyword.value.length == 0) {
		alert("������Ҫ�����Ĺؼ���!");
		frm.keyword.focus();
		return false;
	}

	return true;
}  
</SCRIPT>

  
<TABLE width=766 height="68" border=0 align=center cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR> 
      <TD width="261" background="images/xsmz.gif" bgColor=#ffffff> </TD>
      <TD width="505" bgColor=#ffffff><img src="images/bt_1.gif" width="505" height="75"></TD>
    </TR>
  </TBODY>
</TABLE> 
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=766>
  <TBODY>
  <TR>
    <TD background=images/dhbg2.gif height=25>
      <TABLE width="100%" border=0 align="center" cellPadding=0 cellSpacing=0>
        <TBODY> 
        <TR class=wz6> 
          <TD height=29 width=63> 
            <DIV align=center><A href="index.asp">�� ҳ</A></DIV>
          </TD>
          <TD width=10> 
            <DIV align=center><IMG height=33 src="images/dhtp2.gif" 
            width=2></DIV>
          </TD>
          <TD width=63> 
            <DIV align=center><A 
            href="introduce.asp">�����ſ�</A></DIV>
          </TD>
          <TD width=10> 
            <DIV align=center><IMG height=33 src="images/dhtp2.gif" 
            width=2></DIV>
          </TD>
         
          <TD width=66> 
            <DIV align=center><a href="listallpy.asp">���߷���</a></DIV>
          </TD>
          <TD width=10> 
            <DIV align=center><IMG height=33 src="images/dhtp2.gif" 
            width=2></DIV>
          </TD>
          <TD width=67> 
            <DIV align=center><A 
            href="listworknet.asp">���ϰ���</A></DIV>
          </TD>
          <TD width=10> 
            <DIV align=center><IMG height=33 src="images/dhtp2.gif" 
            width=2></DIV>
          </TD>
          <TD width=66> 
            <DIV align=center><A 
            href="/listallns.asp?act=times">��������</A></DIV>
          </TD>
          <TD width=10> 
            <DIV align=center><IMG height=33 src="images/dhtp2.gif" 
            width=2></DIV>
          </TD>
          <TD width=69> 
            <DIV align=center><A 
            href="/listallvideo.asp">��Ƶ�ļ�</A></DIV>
          </TD>
          <TD width=10> 
            <DIV align=center><IMG height=33 src="images/dhtp2.gif" 
            width=2></DIV>
          </TD>
          <TD width=69> 
            <DIV align=center><A 
            href="consultation.asp">������ѯ</A></DIV>
          </TD>
          <TD width=10><IMG height=33 src="images/dhtp2.gif" 
          width=2></TD>
       
          
          <TD width=95>
            <div align="center"><a href="admingly/admin_index.asp">��̨����</a></div>
          </TD>
        </TR>
        </TBODY>
      </TABLE>
    </TD></TR></TBODY></TABLE>
<%  sub mzfoot()
response.write"<TABLE align=center border=0 cellPadding=0 cellSpacing=0 height=47 width=766><TBODY><TR>"&_
"<TD background=images/bg1.gif class=unnamed6 height=47 vAlign=top><TABLE border=0 cellPadding=0 cellSpacing=0 width='100%'>"&_
"<TBODY><TR><TD height=33 align=center> <SPAN class=wz8>��Ȩ���У������� ������λ�����м�����������޹�˾</SPAN></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></body></html>"
   end sub
%>	