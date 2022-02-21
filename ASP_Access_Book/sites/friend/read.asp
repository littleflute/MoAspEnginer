<%
Option Explicit
dim connpic
dim rs_lar,rspic,rs_apply,rs_friend
dim user_id,sql,willstr
dim picid,cur,pics
%>
<!--#include file="conn.asp"-->
<%
'叛断Session变量是否超时
if isempty(session("user_id")) or session("user_id")="" then
   response.redirect "timeout.htm"
end if
user_id=request("user_id")

Set rs_lar = Server.CreateObject("ADODB.Recordset")
sql="select * from larchives where user_id=" & user_id
rs_lar.open sql,conn,3,2

Set connpic = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("data/picture.mdb")
connpic.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rspic = Server.CreateObject("ADODB.Recordset")
sql="select * from pic where user_id=" & user_id
rspic.open sql,connpic,1,1

set rs_apply=server.createobject("adodb.recordset")
sql="select * from apply where for_id=" & user_id & " and user_id=" & session("user_id")
rs_apply.open sql,conn,1,1
if not(rs_apply.eof and rs_apply.bof) then
	willstr=rs_lar("netname") & "已向您发出交友请求"
end if
rs_apply.close
set rs_apply=nothing

set rs_friend=server.createobject("adodb.recordset")
sql="select * from friend where (for_id=" & session("user_id") & " and user_id=" & user_id & ") or (for_id=" & user_id & " and user_id=" & session("user_id") & ")"
rs_friend.open sql,conn,1,1
if not(rs_friend.eof and rs_friend.bof) then
	willstr=rs_lar("netname") & "是您的好友"
end if
rs_friend.close
set rs_friend=nothing

set rs_apply=server.createobject("adodb.recordset")
sql="select * from apply where user_id =" & session("user_id") & " and for_id=" & user_id
rs_apply.open sql,conn,1,1
if not(rs_apply.eof and rs_apply.bof) then
	willstr="您已向“" & rs_lar("netname") & "”发出交友请求，请静候佳音！"
end if
rs_apply.close
set rs_apply=nothing


if rspic.eof and rspic.bof then
   picid=1
   cur=1
else
   rspic.pagesize=1
   cur=request("cur")
   if cur="" or clng(cur)<1 then cur=1
   if clng(cur)>rspic.pagecount then cur=rspic.pagecount
   rspic.absolutepage=cur
   picid=rspic("id")
end if
   pics=rspic.recordcount
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Namo WebEditor v4.0(Trial)">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>个人档案</title>
<style type="text/css">
<!--
body {
	background-image: url(bihaibg.jpg);
}
-->
</style>
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v3.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}

function MM_showHideLayers() { //v3.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v='hide')?'hidden':v; }
    obj.visibility=v; }
}
//-->
</script>
<link rel="stylesheet" href="ysb.css">
</head>
<body leftmargin="5" topmargin="0">
<table width="772" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="center"><img src="images/top.gif" width="772" height="112" usemap="#Map" border="0"></div></td>
  </tr>
</table>
<p>
<div id="layer2" style="width:765px; height:216px; position:absolute; left:71px; top:130px; z-index:1; visibility:hidden;"> 
  <table border="0" width="761" height="216" cellspacing="0" cellpadding="0">
    <tr bordercolor="#FFFFFF"> 
      <td width="624" valign="top" align="left"> 　 
        <div align="center"> 
          <center>
            <table border="0" width="122%" height="181" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="100%" height="151" valign="top" align="center"><a href="" onClick="return false" ><img name="big" src="display.asp?id=<%=picid%>" onClick="MM_showHideLayers('layer2','','hide');MM_showHideLayers('layer','','show');MM_showHideLayers('layer1','','show');return false;" border="0"></a> 
                </td>
              </tr>
              <tr> 
                <td width="100%" height="12" valign="top" align="center">　</td>
              </tr>
            </table>
          </center>
        </div>
      </td>
    </tr>
  </table>
</div>

  <!------------------------------------------------------------------------------------------------------------------------------> 
  <!---------------------------------------------------------------------------------------------> 
  <table border="0" width="750" height="216" cellspacing="0" cellpadding="0" align="center">
    <tr>
    <td width="6" height="210" valign="middle"> &nbsp;&nbsp; </td>
    <td width="28" height="210" valign="top" bgcolor="#FFCC99"> <font size="2">&nbsp;&nbsp;&nbsp;&nbsp;</font>
    </td>
    <td width="712" height="216" valign="top" bgcolor="#fff4cf" style="border-right-style: solid; border-right-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#F4F4F4">
      <!------------------------------------------------------------------------------------------>
      <table border="0" width="730" height="15" cellspacing="0" cellpadding="2">
        <tr>
          <td width="21" height="8">&nbsp;</td>
          <td width="678" height="8" colspan="5"></td>
                        <td></td>
        </tr>
        <tr>
          <td width="21" height="6"></td>
          <td width="678" height="6" colspan="5"><font size="3" color="#b00000"><b><u><%=rs_lar("name")%></u></b></font>个人档案<p></td>
                        <td></td>
        </tr>
        <tr>
          <td width="21" height="223" valign="top">___</td>
          <td width="39" height="212" valign="top"> <!--------------------------------------------------------------------------------------------->
            <table border="0" width="100%" cellspacing="0" cellpadding="2">
              <tr>
                <td width="100%"  valign="bottom" align="center">
                  <%if pics<>0 then%><a href="read.asp?user_id=<%=user_id%>&cur=<%=cur-1%>" title="查看上一张相片">←</a><%end if%>
                  <%if cur=1 then%> <%if pics=0 then%> <font color="#Fc99d0">还没有上传相片</font>
                  <%else%> <font color="#009900">有相片<font color="red"><%=pics%></font>张</font>
                  <%end if%> <%else%> <font color="#FF0c00">第<font color="red"><%=cur%></font>张</font>
                  <%end if%> <%if pics<>0 then%><a href="read.asp?user_id=<%=user_id%>&cur=<%=cur+1%>" title="查看下一张相片">→</a><%end if%>
                </td>
              </tr>
              <tr>
                  <td width="100%"> <a href="" onClick="return false"><img name="small" border="1" src="display.asp?id=<%=picid%>" width="100" height="135" onClick="MM_showHideLayers('layer2','','show');MM_showHideLayers('layer','','hide');MM_showHideLayers('layer1','','hide');"></a>
                  </td>
              </tr>
              <tr>
                <td width="100%"><font color="black">&nbsp; <%if rs_lar("sex")="女" then%>芳龄<%else%>年龄<%end if%>：<%=rs_lar("age")%></font></td>
              </tr>
              <tr>
                <td width="100%"><font color="black">&nbsp; 学历：<%=rs_lar("education")%></font></td>
              </tr>
              <tr>
                <td width="100%"><font color="black">&nbsp; 星座：<%=rs_lar("star")%></font></td>
              </tr>
              <tr>
                <td width="100%"><font color="black">&nbsp; 人气：<%=rs_lar("renqi")%></font></td>
              </tr>
              <tr>
                <td width="100%"><font color="black">&nbsp; 点击：<%=rs_lar("click")%></font></td>
              </tr>
              <tr>
                <td width="100%" bgcolor="#fff4cf"><font color="black">&nbsp;</font></td>
              </tr>
              <tr>
                <td width="100%" >
              <%if clng(user_id)<>session("user_id") then%>
                  <%if willstr<>"" then%>
                    <p align="center"><%=willstr%>
                  <%else%>
                  <p align="center"><a href="apply.asp?user_id=<%=user_id%>">[发出交友请求]</a>
                  <%end if%>
                </td>
              </tr>
              <tr>
                <td width="100%" bgcolor="#fff4cf">
                  <p align="center"><a href="leaveword.asp?user_id=<%=user_id%>">[写留言]</a>
                </td>
              </tr>
<tr>
                <td width="100%" >
                 
                </td>
              </tr>
              <%else%>
              <tr>
                <td width="100%" bgcolor="#fff4cf" align="center">&nbsp;</td>
              </tr>
              <tr>
                <td width="100%" bgcolor="#fff4cf" align="center"><a href="edit.asp">[修改资料]</a></td>
              </tr>
              <%end if%>
            </table>
            <!--------------------------------------------------------------------------------------------->
          </td>&nbsp;&nbsp;</td>
          <td width="299" height="212" valign="top"> <!--------------------------------------------------------------------------------------------->
            <table border="0" width="100%" cellspacing="0" cellpadding="0" height="237">
              <tr>
                <td width="100%" style="border-bottom-style: solid" bordercolor="#000000" height="16"><b><font color="#FF9900">基本资料</font></b></td>
              </tr>
              <tr>
                <td width="100%" height="6">&nbsp;</td>
              </tr>
              <tr>
                <td width="100%" height="16"><font color="#008000">姓名:&nbsp;<%=rs_lar("name")%></font></td>
              </tr>
              <tr>
                <td width="100%" height="16"><font color="#008000">性别:&nbsp;<%=rs_lar("sex")%></font></td>
              </tr>
              <tr>
                <td width="100%" height="16"><font color="#008000">生日:&nbsp;<%=formatdatetime(rs_lar("britherday"),vblongdate)%></font></td>
              </tr>
              <tr>
                <td width="100%" height="16"><font color="#008000">来自:&nbsp;<%=rs_lar("home")%></font></td>
              </tr>
              <tr>
                <td width="100%" height="16"><font color="#008000">职业:&nbsp;<%=rs_lar("job")%></font></td>
              </tr>
              <tr>
                <td width="100%" height="16"><font color="#008000">单位:&nbsp;<%=rs_lar("company")%></font></td>
              </tr>
              <tr>
                <td width="100%" height="16"><font color="#008000">邮编:&nbsp;<%=rs_lar("postcalcode")%></font></td>
              </tr>
              <tr>
                <td width="100%" height="16"><font color="#008000">电话:&nbsp;<%=rs_lar("tel")%></font></td>
              </tr>
              <tr>
                <td width="100%" height="5"></td>
              </tr>
              <tr>
                <td width="100%" style="border-bottom-style: solid" bordercolor="#000000" height="16"><b><font color="#FF9900">简历&nbsp;</font></b></td>
              </tr>
              <tr>
                <td width="100%" height="10" valign="top">&nbsp;</td>
              </tr>
              <tr>
                <td width="100%" height="69" valign="top"><font color="#008000"><%=rs_lar("fresume")%></font></td>
              </tr>
            </table>
            <!--------------------------------------------------------------------------------------------->
          </td>
          <td width="14" height="212" valign="top"> &nbsp;&nbsp; </td>
          <td width="310" height="212" valign="top"> <!--------------------------------------------------------------------------------------------->
            <table border="0" width="100%" height="29" cellspacing="0" cellpadding="0">
              <tr>
                <td width="270" height="20" style="border-bottom-style: solid" bordercolor="#000000"><font color="#FF9900"><b>个人爱好</b></font></td>
              </tr>
              <tr>
                <td width="268" height="20">&nbsp;</td>
              </tr>
              <tr>
                <td width="268" height="16"><font color="#008000">喜欢的运动:&nbsp;<%=rs_lar("sport")%></font></td>
              </tr>
              <tr>
                <td width="268" height="16"><font color="#008000">喜欢的书籍:&nbsp;<%=rs_lar("book")%></font></td>
              </tr>
              <tr>
                <td width="268" height="16"><font color="#008000">喜欢的音乐:&nbsp;<%=rs_lar("music")%></font></td>
              </tr>
              <tr>
                <td width="268" height="16"><font color="#008000">喜欢的名人:&nbsp;<%=rs_lar("people")%></font></td>
              </tr>
              <tr>
                <td width="268" height="16"><font color="#008000">其它爱好:&nbsp;&nbsp;&nbsp;<%=rs_lar("interest")%></font></td>
              </tr>
              <tr>
                <td width="268" height="16"><font color="#008000">人生格言:&nbsp;&nbsp;&nbsp;<%=rs_lar("adage")%></font></td>
              </tr>
              <tr>
                <td width="268" height="16"><font color="#008000">性格自述:&nbsp;&nbsp;&nbsp;<%=rs_lar("character")%></font></td>
              </tr>
              <tr>
                <td width="268" height="1">&nbsp;&nbsp;</td>
              </tr>
              <tr>
              </tr>
              <tr>
                <td width="270" height="1" style="border-bottom-style: solid" bordercolor="#000000"><b><font color="#FF9900">网上行踪&nbsp;</font></b></td>
              </tr>
              <tr>
                <td width="268" height="1">&nbsp;</td>
              </tr>
              <tr>
                <td width="268" height="20"><font color="#008000">网名: &nbsp;<b><%=rs_lar("netname")%></B></font></td>
              </tr>
              <tr>
                <td width="268" height="20"><font color="#008000">主页: &nbsp;<a href="<%=rs_lar("homepage")%>" target="_blank"><%=rs_lar("homepage")%></font></td>
              </tr>
              <tr>
                <td width="268" height="20"><font color="#008000">Email: &nbsp;<font color="ff0000">只有好友可见</font></font></td>
              </tr>
              <tr>
                <td width="268" height="20"><font color="#008000">网络寻呼机: &nbsp;<%=rs_lar("netcall")%></font></td>
              </tr>
              <tr>
                <td width="268" height="20"><font color="#008000">常去的聊天室:&nbsp; <a href="<%=rs_lar("chatroom")%>" ><%=rs_lar("chatroom")%></font></td>
              </tr>
              <tr>
                <td width="268" height="20"></td>
              </tr>
            </table>
            <!--------------------------------------------------------------------------------------------->
          </td>
          <td width="22" height="212" valign="top"> &nbsp;&nbsp;&nbsp;</td>
        </tr>
        <tr>
          <td width="23" height="1"></td>
          <td width="680" height="1" colspan="5">
            <table border="0" width="100%" cellspacing="0" cellpadding="0">
              <tr>
                <td width="67%"> </td>
                <td width="33%">
                  <p align="right"><font color="#000080">此档案已被浏览了<font color="#ff0000"><b><%=rs_lar("click")%></b></font>次</font>
                </td>
              </tr>
            </table>
          </td>
                        <td></td>
        </tr>
        <tr>
          <td width="23" height="8"></td>
          <td width="680" height="8" colspan="5">&nbsp;</td>
                        <td></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</div></p>
<map name="FPMap2"></map>
<map name="Map">
  <area shape="rect" coords="231,96,275,114" href="default.asp">
  <area shape="rect" coords="294,93,358,111" href="your.asp">
  <area shape="rect" coords="383,93,448,112" href="list.asp">
  <area shape="rect" coords="468,92,532,112" href="register.asp">
  <area shape="rect" coords="554,92,616,112" href="sendphoto.asp">
  <area shape="rect" coords="638,91,704,113" href="pics.asp">
</map>
</body>

</html>
<%
  rs_lar("click")=rs_lar("click")+1
  rs_lar.update
  rs_lar.close
  set rs_lar=nothing
%>
