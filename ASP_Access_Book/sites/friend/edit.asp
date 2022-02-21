<!--#include file="conn.asp"-->
<%
'叛断此用户是否已经入会
if session("user_id")="" then
     response.redirect "notreg.htm"
	 response.end
end if

Set rs_lar = Server.CreateObject("ADODB.Recordset")
sql="select * from larchives where user_id=" & session("user_id")
rs_lar.open sql,conn,3,2
if rs_lar.eof and rs_lar.bof then
   response.redirect "notregist.htm"
   response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Namo WebEditor v4.0(Trial)">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>建立档案</title>
<style>
<!--
-->
</style>
<link rel="stylesheet" href="ysb.css">
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" leftmargin="5" topmargin="0" >
<table width="772" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="images/top.gif" width="772" height="112" usemap="#Map" border="0"></td>
  </tr>
</table>
<map name="FPMap1"> <area shape="rect" coords="381, 4, 446, 18" href="sendphoto.asp"> 
  <area shape="rect" coords="283, 4, 357, 18" href="register.asp"
 target="_blank"> <area shape="rect" coords="190, 4, 264, 18" href="list.asp"> 
  <area shape="rect" coords="8, 2, 71, 16" href="your.asp"> <area shape="rect" coords="102, 4, 168, 18" href="no.asp"> 
</map> 
<table border="0" width="776" cellspacing="0" cellpadding="0">
  <tr>
    <td width="3%">&nbsp; </td>
    <td width="71%">&nbsp; </td>
    <td width="1%">&nbsp; </td>
    <td width="25%"></td>
  </tr>
  <tr>
    <td width="3%" valign="top">
    </td>
    <td width="71%" valign="top"> 
      <form action="eadd.asp" method="POST">
<div align="left">
<table border="0" width="526" cellspacing="0" cellpadding="0">
    <td width="600">
      <table border="0" width="552" cellspacing="0" height="204" cellpadding="0">
                  <tr bgcolor="#067B0F"> 
                    <td height="11" colspan="3"><font color="#FFFFFF"><b>请认真填写您的个人资料</b></font></td>
        </tr>
        <tr>
                    <td height="4" colspan="3"><font color="#FFFFFF">&nbsp;</font></td>
        </tr>
        <tr>
                    <td height="7" style="border-bottom-style: solid" bordercolor="#000000" colspan="3"><b><font color="#FF9900">您的基本资料</font></b></td>
        </tr>
        <tr>
          <td width="119" height="18"><font color="#008000"><b>*真实姓名:</b></font></td>
                    <td width="260" height="18"> 
                      <input type="text" name="name" value="<%=rs_lar("name")%>" size="10" style="border-style: solid; border-width: 1"></td>
                    <td width="173" height="18"><i>2-10字符</i></td>
        </tr>
        <tr>
          <td width="119" height="18"><font color="#008000"><b>*您的性别:</b></font></td>
                    <td width="260" height="18"> 
                      <INPUT <%if rs_lar("sex")="男" then%>CHECKED<%end if%> name=sex type=radio
            value=男> <SPAN class=tt4>男&nbsp; <INPUT <%if rs_lar("sex")="女" then%>CHECKED<%end if%> name=sex type=radio value=女>   
            女&nbsp;</SPAN> </td>   
                    <td width="173" height="18"> </td>   
        </tr>   
        <tr>   
          <td width="119" height="18"><font color="#008000"><b>*您的生日:</b></font></td>   
                    <td width="260" height="18"> 
                      <select size="1" name="Byear">   
            <%for i=1940 to year(date)-3%>   
              <option value="<%=i%>" <%if year(rs_lar("britherday"))=i then%> selected <%end if%>><%=i%></option>   
            <%next%>   
            </select>年<select size="1" name="Bmonth">   
            <%for i=1 to 12%>   
              <option value="<%=i%>" <%if month(rs_lar("britherday"))=i then%> selected <%end if%>><%=i%></option>   
            <%next%>   
            </select>月<select size="1" name="Bday">   
            <%for i=1 to 31%>   
              <option value="<%=i%>" <%if day(rs_lar("britherday"))=i then%> selected <%end if%>><%=i%></option>   
            <%next%>   
            </select>日</td>   
                    <td width="173" height="18"> </td>   
        </tr>   
        <tr>   
          <td width="119" height="11"><font color="#008000"><b>*您的籍贯:</b></font></td>   
                    <td width="260" height="11" valign="middle"> 
                      <input type="text" name="home" size="11" style="border-style: solid; border-width: 1" value="<%=rs_lar("home")%>"></td> 
                    <td width="173" height="11" valign="middle"><i>不超过50个字符</i></td> 
        </tr> 
        <tr> 
          <td width="119" height="11"><font color="#008000"><b>*您的学历</b></font></td> 
                    <td width="260" height="11"><SPAN class=tt4> 
                      <SELECT name=education 
            size=1> <OPTION selected value=4>本科</OPTION> <OPTION 
              value=2>中专</OPTION> <OPTION value=3>大专</OPTION> <OPTION 
              value=1>高中</OPTION> <OPTION value=6>博士</OPTION> <OPTION 
              value=5>硕士</OPTION> <OPTION value=0>在校生</OPTION> <OPTION 
              value=7>其他</OPTION></SELECT> </SPAN></td> 
                    <td width="173" height="11"></td> 
        </tr> 
        <tr> 
          <td width="119" height="19"><font color="#008000"><b>*您的职业:</b></font></td> 
                    <td width="260" height="19"> 
                      <SELECT name="job" size="1"> 
              <OPTION selected value=管理>管理</OPTION> <OPTION value=工程>工程</OPTION>
              <OPTION value=金融财务>金融财务</OPTION> <OPTION value=技术>技术</OPTION> <OPTION
              value=经济业务>经济业务</OPTION> <OPTION value=法律>法律</OPTION> <OPTION
              value=保险>保险</OPTION> <OPTION value=教师>教师</OPTION> <OPTION
              value=科研>科研</OPTION> <OPTION value=设计>设计</OPTION> <OPTION
              value=行政>行政</OPTION> <OPTION value=文体卫生>文体卫生</OPTION> <OPTION
              value=服务>服务</OPTION> <OPTION value=军队公安>军队公安</OPTION> <OPTION
              value=IT>IT</OPTION> <OPTION value=公务员>公务员</OPTION> <OPTION
              value=其它>其它</OPTION></SELECT> </td>
                    <td width="173" height="19"> </td>
        </tr>
        <tr>
          <td width="119" height="18"><font color="#008000"><b>&nbsp;</b></font><b>您的单位:</b></td>
                    <td width="260" height="18"> 
                      <input type="text" name="company" size="31" style="border-style: solid; border-width: 1" value="<%=rs_lar("company")%>"></td>
                    <td width="173" height="18"><i>不超过50个字符</i></td>
        </tr>
        <tr>
          <td width="119" height="18"><font color="#008000"><b>*您的邮编:</b></font></td>
                    <td width="260" height="18"> 
                      <input type="text" name="postcalcode" size="6" style="border-style: solid; border-width: 1" value="<%=rs_lar("postcalcode")%>"></td>
                    <td width="173" height="18"><i>6个数字</i></td>
        </tr>
        <tr>
          <td width="119" height="14"><font color="#008000"><b>*您的电话:</b></font></td>
                    <td width="260" height="14"> 
                      <input type="text" name="tel" size="15" style="border-style: solid; border-width: 1" value="<%=rs_lar("tel")%>"></td>
                    <td width="173" height="14"><i>不超过20个字符</i></td>
        </tr>
        <tr>
          <td width="119" height="11"></td>
                    <td width="260" height="11"></td>
                    <td width="173" height="11"></td>
        </tr>
        <tr>
                    <td height="14" colspan="3"> 
                      <table border="0" width="100%" cellspacing="0" cellpadding="0">
              <tr>
                <td width="100%" style="border-bottom-style: solid" bordercolor="#000000"><b><font color="#FF9900">您的个人简历</font></b></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="119" height="16">&nbsp;个人简历:</td>
                    <td width="260" height="54" rowspan="2"> 
                      <textarea rows="3" name="fresume" cols="37" wrap="hard" style="border-style: solid; border-width: 1"><%=rs_lar("fresume")%></textarea></td>
                    <td width="173" height="19"><i>100字以内</i></td>
        </tr>
        <tr>
          <td width="119" height="37"></td>
                    <td width="173" height="44"></td>
        </tr>
        <tr>
          <td width="119" height="10"></td>
                    <td width="260" height="10"></td>
                    <td width="173" height="10"></td>
        </tr>
        <tr>
                    <td height="5" colspan="3"> 
                      <table border="0" width="100%" cellspacing="0" cellpadding="0">
              <tr>
                <td width="100%" style="border-bottom-style: solid" bordercolor="#000000"><b><font color="#FF9900">您的网上行踪</font></b></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="119" height="5"><font color="#008000"><b>*您的网名:</b></font></td>
                    <td width="260" height="5"> 
                      <input disabled type="text" name="netname" size="11" style="border-style: solid; border-width: 1" value="<%=rs_lar("netname")%>"></td>
                    <td width="173" height="5"><i>2-10字符</i></td>
        </tr>
        <tr>
          <td width="119" height="8"><font color="#008000"><b>&nbsp;</b></font><b>个人主页:</b></td>
                    <td width="260" height="8"> 
                      <input type="text" name="homepage" size="30" style="border-style: solid; border-width: 1" value="<%=rs_lar("homepage")%>"></td>
                    <td width="173" height="8">要填写<font color="#FF0000"><b>http://</b></font></td>
        </tr>
        <tr>
          <td width="119" height="5"><font color="#008000"><b>*电子邮件:</b></font></td>
                    <td width="260" height="5"> 
                      <input type="text" name="email" size="20" style="border-style: solid; border-width: 1" value="<%=rs_lar("email")%>"></td>
                    <td width="173" height="5"><i>一定要真实否则收不到网友信息</i></td>
        </tr>
        <tr>
          <td width="119" height="3"><font color="#008000"><b>&nbsp;</b></font><b>网络寻呼机:</b></td>
                    <td width="260" height="3"> 
                      <input type="text" name="netcall" size="14" style="border-style: solid; border-width: 1" value="<%=rs_lar("netcall")%>">(网络寻呼机号码）</td>
                    <td width="173" height="3"><i>不超过12个字符</i></td>
        </tr>
        <tr>
          <td width="119" height="16"><font color="#008000"><b>&nbsp;</b></font><b>常去的聊天室:</b></td>
                    <td width="260" height="16"> 
                      <input type="text" name="chatroom" size="20" style="border-style: solid; border-width: 1 value=" <%=rs_lar("chatroom")%>""></td>
                    <td width="173" height="16"><i>不超过50个字符</i></td>
        </tr>
        <tr>
          <td width="119" height="11"></td>
                    <td width="260" height="11"></td>
                    <td width="173" height="11"></td>
        </tr>
        <tr>
                    <td height="4" colspan="3"> 
                      <table border="0" width="100%" cellspacing="0" cellpadding="0">
              <tr>
                <td width="100%" style="border-bottom-style: solid" bordercolor="#000000"><b><font color="#FF9900">您的性格与爱好</font></b></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="119" height="5"><font color="#008000"><b>*喜欢的运动:</b></font></td>
                    <td width="260" height="5"> 
                      <input type="text" name="sport" size="20" style="border-style: solid; border-width: 1" value="<%=rs_lar("sport")%>"></td>
                    <td width="173" height="5"><i>不超过30个字符</i></td>
        </tr>
        <tr>
          <td width="119" height="7"><font color="#008000"><b>*喜欢的书籍:</b></font></td>
                    <td width="260" height="7"> 
                      <input type="text" name="book" size="20" style="border-style: solid; border-width: 1" value="<%=rs_lar("book")%>"></td>
                    <td width="173" height="7"><i>不超过50个字符</i></td>
        </tr>
        <tr>
          <td width="119" height="6"><font color="#008000"><b>*喜欢的音乐:</b></font></td>
                    <td width="260" height="6"> 
                      <input type="text" name="music" size="25" style="border-style: solid; border-width: 1" value="<%=rs_lar("music")%>"></td>
                    <td width="173" height="6"><i>不超过50个字符</i></td>
        </tr>
        <tr>
          <td width="119" height="5"><font color="#008000"><b>*喜欢的名人:</b></font></td>
                    <td width="260" height="5"> 
                      <input type="text" name="people" size="20" style="border-style: solid; border-width: 1" value="<%=rs_lar("people")%>"></td>
                    <td width="173" height="5"><i>不超过30个字符</i></td>
        </tr>
        <tr>
          <td width="119" height="5"><font color="#008000"><b>&nbsp;</b></font><b>其它爱好或特长:</b></td>
                    <td width="260" height="5"> 
                      <input type="text" name="interest" size="30" style="border-style: solid; border-width: 1" value="<%=rs_lar("interest")%>"></td>
                    <td width="173" height="5"><i>不超过50个字符</i></td>
        </tr>
        <tr>
          <td width="119" height="11"><font color="#008000"><b>*人生格言:</b></font></td>
                    <td width="260" height="11"> 
                      <input type="text" name="adage" size="31" style="border-style: solid; border-width: 1" value="<%=rs_lar("adage")%>"></td>
                    <td width="173" height="11"><i>不超过50个字符</i></td>
        </tr>
        <tr>
          <td width="119" height="9"><font color="#008000"><b>*性格自述:</b></font></td>
                    <td width="260" height="9"> 
                      <input type="text" name="character" size="31" style="border-style: solid; border-width: 1" value="<%=rs_lar("character")%>"></td>
                    <td width="173" height="9"><i>不超过50个字符</i></td>
        </tr>
        <tr>
          <td width="119" height="16">　</td>
                    <td width="260" height="16"> 
                      <p align="right"></td>
                    <td width="173" height="16"> 
                      <p align="right"><input type="submit" value="提交" name="B1"><input type="reset" value="重写" name="B2">
          </td>
        </tr>
      </table>
    </td>
</table>
</div>
</form>


    </td>
    <td width="1%" valign="top">

    </td>
    <td width="25%" valign="top"> 
      <table border="1" width="100%" bordercolor="#CC3300" cellspacing="0" cellpadding="0" height="707">
        <tr bgcolor="#067B0F"> 
          <td width="100%" height="13"> 
            <p align="center"><font color="#FFFFFF"><b>给网友的话</b></font></p>
        </td>
      </tr>
      <tr>
        <td width="100%" height="690" valign="top">
            <table border="0" width="100%" height="162" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="100%" height="241" valign="top">&nbsp;&nbsp; 为了体现您完美与真实的自我，请祥细和认真的填写您的个人档案。 
                  <p>&nbsp;&nbsp; 如果您想上传照片，请提交了个人档案以后，进入<a href="sendphoto.asp">[相片上传]</a>栏目。</p>
                  <p>&nbsp;&nbsp; 本交友中心初建于2001年4月，至今已得到了很多网友的支持。在此特向所有支持我们网站的网友表示感谢！</p>
                  <p>&nbsp;&nbsp;&nbsp; 祝愿所有的网友都能找到人生知己。&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                  </p>
                </td>
              </tr>
              <tr> 
                <td width="100%" height="1" valign="top"></td>
              </tr>
            </table>
        </td>
      </tr>
    </table>

    </td>
  </tr>
</table>
<map name="Map"> 
  <area shape="rect" coords="226,94,278,110" href="default.asp">
  <area shape="rect" coords="291,93,360,112" href="your.asp">
  <area shape="rect" coords="381,93,446,112" href="list.asp">
  <area shape="rect" coords="469,92,532,113" href="register.asp">
  <area shape="rect" coords="554,91,618,111" href="sendphoto.asp">
  <area shape="rect" coords="640,92,704,112" href="pics.asp">
</map>
</body>

</html>
<%
rs_lar.close
set rs_lar=nothing
set conn=nothing
%>