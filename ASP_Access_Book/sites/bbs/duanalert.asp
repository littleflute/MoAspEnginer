<% If session("AdminUID")<>"" Then %>
<!--#include file="conn.asp"-->
 <%
uer=session("AdminUID")
sql_news5="Select count(*) as numb From duanx where [to]="&"'"&uer&"' and  flag  like  '否' "
set rs_news5=conn.execute(sql_news5)
if rs_news5("numb")>0 then
 %>
 <script LANGUAGE="JavaScript">
 alert(" 您有新的消息!");
</script>
<% End If %>
<a href=javascript:openScript('duan.asp',450,350)><font size="-1"><img src="pic/msg_no_new_bar.gif" width="16" height="16" border="0"><font color="#000000">收件箱：</font></font><FONT COLOR=#ff0000 size="-1"> 
<%  =rs_news5("numb") %>
新 </FONT></a> <font size="-1">
<% 
rs_news5.close
%>
<script language="javascript">
function openUser(id) {
	var Win = window.open("dispuser.asp?name="+id,"openScript","width=350,height=300,resizable=1,scrollbars=1,menubar=0,status=1" );
}

function openScript(url, width, height){
	var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}

function openDis(bid,rid,id){
	self.location="dispbbs.asp?boardid="+bid+"&RootID="+rid+"&id="+id
}

function PopWindow()
{ 
openScript('messanger.asp?action=newmsg',420,320); 
}
var nn = !!document.layers;
var ie = !!document.all;

if (nn) {
  netscape.security.PrivilegeManager.enablePrivilege("UniversalSystemClipboardAccess");
  var fr=new java.awt.Frame();
  var Zwischenablage = fr.getToolkit().getSystemClipboard();
}

function submitonce(theform){
//if IE 4+ or NS 6+
if (document.all||document.getElementById){
//screen thru every element in the form, and hunt down "submit" and "reset"
for (i=0;i<theform.length;i++){
var tempobj=theform.elements[i]
if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
//disable em
tempobj.disabled=true
}
}
}
</script>
</font>
<% End If %>