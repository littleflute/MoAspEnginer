<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>学工论坛--短消息</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="duanx.files/forum.css" rel=stylesheet>
<META content="MSHTML 5.00.3315.2870" name=GENERATOR></HEAD>
<BODY aLink=#333333 bgColor=#eeeeee link=#333333 
onkeydown="if(event.keyCode==13 &amp;&amp; event.ctrlKey)messager.submit()" 
topMargin=10 vLink=#333333>
<FORM action=duanaddsave.asp method=post name=form>
  <TABLE align=center bgColor=#009933 border=0 cellPadding=0 cellSpacing=0 
width="95%">
    <TBODY> 
    <TR> 
      <TD> 
        <TABLE border=0 cellPadding=3 cellSpacing=1 width="100%">
          <TBODY> 
          <TR> 
            <TD align=middle bgColor=#ccff99 colSpan=3><FONT color=#333333 
            face=宋体><B>发送短消息</B></FONT></TD>
          </TR>
          <TR> 
                <TD align=middle bgColor=#eeeeee colSpan=3 vAlign=center><a href="duan.asp"><IMG 
            alt=收件箱 border=0 src="duanx.files/inboxpm.gif"></a> &nbsp; <A 
            href="duanadd.asp"><IMG 
            alt=发送消息 border=0 src="duanx.files/newpm.gif"></A></TD>
          </TR>
          <TR> 
            <TD align=middle bgColor=#ccff99 colSpan=2> 
              <INPUT name=action 
            type=hidden value=send>
              <FONT color=#333333 
            face=宋体><B>请完整输入下列信息</B></FONT></TD>
          </TR>
          <TR> 
            <TD bgColor=#eeeeee vAlign=center><FONT color=#333333 
            face=宋体><B>收件人：</B></FONT></TD>
            <TD bgColor=#eeeeee vAlign=center> 
              <INPUT name=to>
            </TD>
          </TR>
          <TR> 
            <TD bgColor=#eeeeee vAlign=top width="30%"><FONT color=#333333 
            face=宋体><B>标题：</B></FONT></TD>
            <TD bgColor=#eeeeee vAlign=center> 
              <INPUT maxLength=80 name=title 
            size=36>
              <input name="from" type="hidden" id="user_name" value="<%=session("AdminUID")%>">
            </TD>
          </TR>
          <TR> 
            <TD bgColor=#eeeeee vAlign=top width="30%"><FONT color=#333333 
            face=宋体><B>内容：</B><BR>
              Ctrl+Enter发送</FONT></TD>
            <TD bgColor=#eeeeee vAlign=center> 
              <TEXTAREA cols=35 name=content rows=6 title=Ctrl+Enter发送></TEXTAREA>
            </TD>
          </TR>
          <TR> 
            <TD align=middle bgColor=#ccff99 colSpan=2 vAlign=center> 
              <INPUT name='Submit"' type=submit value="发 送">
              &nbsp; 
              <INPUT name=Clear type=reset value="清 除">
            </TD>
          </TR>
          </TBODY> 
        </TABLE>
      </TD>
    </TR>
    </TBODY> 
  </TABLE>
</FORM>
<P> 
  <SCRIPT language=javascript>
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
</SCRIPT>
</BODY></HTML>
