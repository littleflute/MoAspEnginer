<!--#include file="conn.asp"-->
<%
dim rs_lar
dim sql
dim i
'叛断Session变量是否超时
if isnull(session("user_id")) then
   response.redirect "timeout.htm"
end if
'叛断此用户是否已经入会
if session("user_id")="1" then
    response.redirect "notreg.htm"
	 response.end
end if
'判断是否已经填写档案
Set rs_lar = Server.CreateObject("ADODB.Recordset")
sql="select * from larchives where user_id =" & session("user_id")
rs_lar.open sql,conn,3,2
if not(rs_lar.eof and rs_lar.bof) then
   response.redirect "haveregist.htm"
   response.end
end if
rs_lar.close
set rs_lar=nothing
set conn=nothing
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>建立档案</title>
<style>
<!--
-->
</style>
<link rel="stylesheet" href="ysb.css">
</head>

<body leftmargin="5" topmargin="0">
<table width="772" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="images/top.gif" width="772" height="112" usemap="#Map" border="0"></td>
  </tr>
</table>
<table border="0" width="776" cellspacing="0" cellpadding="0">
  <tr align="center"> 
    <td style="border-bottom-style: solid" width="136" height="19"><a href="default.asp" style="color:#ffffff"><font size="2"><font color="#000000">交友首页</font></font></a></td>
    <td style="border-bottom-style: solid" width="126" height="19"><a href="reg.asp" class="x">申请入会</a></td>
    <td style="border-bottom-style: solid" width="126" height="19"><a href="register.asp" class="x">档案注册</a></td>
    <td style="border-bottom-style: solid" width="126" height="19"><a href="your.asp" class="x">个人专栏</a></td>
    <td style="border-bottom-style: solid" width="126" height="19"><a href="list.asp" class="x">网友总览</a></td>
    <td style="border-bottom-style: solid" width="117" height="19"><a href="sendphoto.asp" class="x">相片上传</a></td>
    <td style="border-bottom-style: solid" width="9" height="19">&nbsp;</td>
  </tr>
</table>
<table border="0" width="776" cellspacing="0" cellpadding="0" bgcolor="#FFEBE1">
  <tr> 
    <td width="552" valign="top"> 
      <form action="ladd.asp" method="POST">
        <div align="left"> 
          <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr valign="top"> 
              <td width="600"> 
                <table border="0" width="100%" cellspacing="0" cellpadding="0">
                  <tr bgcolor="#067B0F" align="center"> 
                    <td height="19" colspan="3"><font color="#FFFFFF"><b>请认真填写您的个人资料</b><font color="#FFCCCC">（带*号的必须填写） 
                      </font></font></td>
                </tr>
                <tr> 
                  <td height="4" colspan="3"><font color="#FFFFFF">&nbsp;</font></td>
                </tr>
                <tr> 
                  <td height="7" style="border-bottom-style: solid" bordercolor="#000000" colspan="3"><b><font color="#FF9900">您的基本资料</font></b></td>
                </tr>
                <tr> 
                  <td width="119" height="18"><font color="#008000"><b>*真实姓名:</b></font></td>
                  <td width="294" height="18"> 
                    <input type="text" name="name" size="10" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="18"><i>2-10字符</i></td>
                </tr>
                <tr> 
                  <td width="119" height="18"><font color="#008000"><b>*您的性别:</b></font></td>
                  <td width="294" height="18"> 
                    <INPUT CHECKED name=sex type=radio
            value=男>
                    <SPAN class=tt4>男&nbsp; 
                    <INPUT name=sex type=radio value=女>
                    女&nbsp;</SPAN> </td>
                  <td width="139" height="18"> </td>
                </tr>
                <tr> 
                  <td width="119" height="18"><font color="#008000"><b>*您的生日:</b></font></td>
                  <td width="294" height="18"> 
                    <select size="1" name="Byear">
                      <%for i=1940 to year(date)-3%> 
                      <option value="<%=i%>"><%=i%></option>
                      <%next%> 
                    </select>
                    年 
                    <select size="1" name="Bmonth">
                      <%for i=1 to 12%> 
                      <option value="<%=i%>"><%=i%></option>
                      <%next%> 
                    </select>
                    月 
                    <select size="1" name="Bday">
                      <%for i=1 to 31%> 
                      <option value="<%=i%>"><%=i%></option>
                      <%next%> 
                    </select>
                    日</td>
                  <td width="139" height="18"> </td>
                </tr>
                <tr> 
                  <td width="119" height="11"><font color="#008000"><b>*您的籍贯:</b></font></td>
                  <td width="294" height="11" valign="middle"><SPAN class=tt4> 
                    <SELECT class=small13
            name=province size=1>
                      <option value="上海市">上海市</option>
                      <option value="北京市">北京市</option>
                      <option value="天津市">天津市</option>
                      <option value="河北省">河北省</option>
                      <option value="山西省">山西省</option>
                     
                      <option value="辽宁省">辽宁省</option>
                      <option value="吉林省">吉林省</option>
                      <option value="黑龙江省">黑龙江省</option>
                      <option value="江苏省">江苏省</option>
                      <option value="浙江省">浙江省</option>
                      <option value="安徽省">安徽省</option>
                      <option value="福建省">福建省</option>
                      <option value="江西省">江西省</option>
                      <option value="山东省">山东省</option>
                      <option value="河南省">河南省</option>
                      <option value="湖北省">湖北省</option>
                      <option value="湖南省" >湖南省</option>
                      
                      <option value="四川省">四川省</option>
                      <option value="贵州省">贵州省</option>
                      <option value="云南省">云南省</option>
                     
                      <option value="陕西省">陕西省</option>
                      <option value="甘肃省">甘肃省</option>
                      <option value="青海省">青海省</option>
                     
                      <option value="重庆市">重庆市</option>
                      <option value="海南省">海南省</option>
                      <option value="国外">国外</option>
                      <option value="其他">其他</option>
                      <option value="深圳市">深圳市</option>
                      <option value="广东省" selected>广东省</option>
                    </SELECT>
                    </SPAN> 
                    <input type="text" name="home" size="11" style="border-style: solid; border-width: 1">
                    (市/区）</td>
                  <td width="139" height="11" valign="middle"><i>(市/区)2-10字符</i></td>
                </tr>
                <tr> 
                  <td width="119" height="11"><font color="#008000"><b>*您的学历</b></font></td>
                  <td width="294" height="11"><SPAN class=tt4> 
                    <SELECT name=education
            size=1>
                      <OPTION selected value=本科>本科</OPTION>
                      <OPTION
              value=中专>中专</OPTION>
                      <OPTION value=大专>大专</OPTION>
                      <OPTION
              value=高中>高中</OPTION>
                      <OPTION value=博士>博士</OPTION>
                      <OPTION
              value=硕士>硕士</OPTION>
                      <OPTION value=在校生>在校生</OPTION>
                      <OPTION
              value=其他>其他</OPTION>
                    </SELECT>
                    </SPAN></td>
                  <td width="139" height="11"></td>
                </tr>
                <tr> 
                  <td width="119" height="19"><font color="#008000"><b>*您的职业:</b></font></td>
                  <td width="294" height="19"> 
                    <SELECT name="job" size="1">
                      <OPTION selected value=管理>管理</OPTION>
                      <OPTION value=工程>工程</OPTION>
                      <OPTION value=金融财务>金融财务</OPTION>
                      <OPTION value=技术>技术</OPTION>
                      <OPTION
              value=经济业务>经济业务</OPTION>
                      <OPTION value=法律>法律</OPTION>
                      <OPTION
              value=保险>保险</OPTION>
                      <OPTION value=教师>教师</OPTION>
                      <OPTION
              value=科研>科研</OPTION>
                      <OPTION value=设计>设计</OPTION>
                      <OPTION
              value=行政>行政</OPTION>
                      <OPTION value=文体卫生>文体卫生</OPTION>
                      <OPTION
              value=服务>服务</OPTION>
                      <OPTION value=军队公安>军队公安</OPTION>
                      <OPTION
              value=IT>IT</OPTION>
                      <OPTION value=公务员>公务员</OPTION>
                      <OPTION
              value=其它>其它</OPTION>
                    </SELECT>
                  </td>
                  <td width="139" height="19"> </td>
                </tr>
                <tr> 
                  <td width="119" height="18"><font color="#008000"><b>&nbsp;</b></font><b>您的单位:</b></td>
                  <td width="294" height="18"> 
                    <input type="text" name="company" size="31" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="18"><i>不超过50个字符</i></td>
                </tr>
                <tr> 
                  <td width="119" height="18"><font color="#008000"><b>*您的邮编:</b></font></td>
                  <td width="294" height="18"> 
                    <input type="text" name="postcalcode" size="6" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="18"><i>6个数字</i></td>
                </tr>
                <tr> 
                  <td width="119" height="14"><font color="#008000"><b>*您的电话:</b></font></td>
                    <td width="294" height="14"> 
                      <input type="text" name="tel" size="15" style="border-style: solid; border-width: 1">
                      （必须是不少于7位的数字） </td>
                  <td width="139" height="14"><i>不超过20个字符</i></td>
                </tr>
                <tr> 
                  <td width="119" height="11"></td>
                  <td width="294" height="11"></td>
                  <td width="139" height="11"></td>
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
                  <td width="119" height="20">&nbsp;个人简历:</td>
                  <td width="294" height="1" rowspan="3"> 
                    <textarea rows="4" name="fresume" cols="29" wrap="hard" style="border-style: solid; border-width: 1"></textarea>
                  </td>
                  <td width="139" height="1" rowspan="2"><i>200字以内</i></td>
                </tr>
                <tr> 
                  <td width="119" height="34"></td>
                </tr>
                <tr> 
                  <td width="119" height="1"></td>
                  <td width="139" height="1"></td>
                </tr>
                <tr> 
                  <td width="119" height="10"></td>
                  <td width="294" height="10"></td>
                  <td width="139" height="10"></td>
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
                  <td width="294" height="5"> 
                    <input type="text" name="netname" size="11" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="5"><i>2-10字符</i></td>
                </tr>
                <tr> 
                  <td width="119" height="8"><font color="#008000"><b>&nbsp;</b></font><b>个人主页:</b></td>
                  <td width="294" height="8"> 
                    <input type="text" name="homepage" size="30" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="8"><i>不超过50个字符</i></td>
                </tr>
                <tr> 
                  <td width="119" height="5"><font color="#008000"><b>*电子邮件:</b></font></td>
                  <td width="294" height="5"> 
                    <input type="text" name="email" size="20" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="5"><i>不超过50个字符</i></td>
                </tr>
                <tr> 
                  <td width="119" height="3"><font color="#008000"><b>&nbsp;</b></font><b>网络寻呼机:</b></td>
                  <td width="294" height="3"> 
                    <select size="1" name="netcall">
                      <option selected value="-没有-">-没有-</option>
                      <option value="ICQ">ICQ</option>
                      <option value="OICQ">OICQ</option>
                    </select>
                    <input type="text" name="netcallcode" size="14" style="border-style: solid; border-width: 1">
                    (网络寻呼机号码）</td>
                  <td width="139" height="3"><i>不超过12个字符</i></td>
                </tr>
                <tr> 
                  <td width="119" height="16"><font color="#008000"><b>&nbsp;</b></font><b>常去的聊天室:</b></td>
                  <td width="294" height="16"> 
                    <input type="text" name="chatroom" size="20" style="border-style: solid; border-width: 1">
                    （请写具体网址）</td>
                  <td width="139" height="16"><i>不超过50个字符</i></td>
                </tr>
                <tr> 
                  <td width="119" height="11"></td>
                  <td width="294" height="11"></td>
                  <td width="139" height="11"></td>
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
                  <td width="294" height="5"> 
                    <input type="text" name="sport" size="20" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="5"><i>不超过30个字符</i></td>
                </tr>
                <tr> 
                  <td width="119" height="7"><font color="#008000"><b>*喜欢的书籍:</b></font></td>
                  <td width="294" height="7"> 
                    <input type="text" name="book" size="20" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="7"><i>不超过50个字符</i></td>
                </tr>
                <tr> 
                  <td width="119" height="6"><font color="#008000"><b>*喜欢的音乐:</b></font></td>
                  <td width="294" height="6"> 
                    <input type="text" name="music" size="25" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="6"><i>不超过50个字符</i></td>
                </tr>
                <tr> 
                  <td width="119" height="5"><font color="#008000"><b>*喜欢的名人:</b></font></td>
                  <td width="294" height="5"> 
                    <input type="text" name="people" size="20" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="5"><i>不超过30个字符</i></td>
                </tr>
                <tr> 
                  <td width="119" height="5"><font color="#008000"><b>&nbsp;</b></font><b>其它爱好或特长:</b></td>
                  <td width="294" height="5"> 
                    <input type="text" name="interest" size="30" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="5"><i>不超过50个字符</i></td>
                </tr>
                <tr> 
                  <td width="119" height="11"><font color="#008000"><b>*人生格言:</b></font></td>
                  <td width="294" height="11"> 
                    <input type="text" name="adage" size="31" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="11"><i>不超过50个字符</i></td>
                </tr>
                <tr> 
                  <td width="119" height="9"><font color="#008000"><b>*性格自述:</b></font></td>
                  <td width="294" height="9"> 
                    <input type="text" name="character" size="31" style="border-style: solid; border-width: 1">
                  </td>
                  <td width="139" height="9"><i>不超过50个字符</i></td>
                </tr>
                <tr> 
                  <td width="119" height="16">　</td>
                  <td width="294" height="16"> 
                    <p align="right">
                  </td>
                  <td width="139" height="16"> 
                    <p align="right"> 
                      <input type="submit" value="提交" name="B1">
                      <input type="reset" value="重写" name="B2">
                  </td>
                </tr>
              </table>
              </td>
            
          </table>
        </div>
      </form>
    </td>
    <td width="10" valign="top"> </td>
    <td width="192" valign="top"> 
      <table border="1" width="100%" bordercolor="#CC0000" cellspacing="0" cellpadding="0" height="707">
        <tr bgcolor="#067B0F"> 
          <td width="100%" height="13"> 
            <p align="center"><font color="#FFFFFF"><b><a href="sendphoto.asp"><font color="#FFCCCC">上传照片</font></a></b></font></p>
          </td>
        </tr>
        <tr> 
          <td width="100%" height="690" valign="top"> 
            <table border="0" width="100%" height="162" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="100%" height="241" valign="top">&nbsp;&nbsp; <br>
                  为了体现您完美与真实的自我，请祥细和认真的填写您的个人档案。 
                  <p>&nbsp;&nbsp; 如果您想上传照片，请提交了个人档案以后，进入<a href="sendphoto.asp">[相片上传]</a>栏目。</p>
                  <p>&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </p>
                </td>
              </tr>
              <tr> 
                <td width="100%" height="17" valign="top"> 
                  <p align="right"> 
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
  <area shape="rect" coords="233,94,273,113" href="default.asp">
  <area shape="rect" coords="293,92,354,110" href="your.asp">
  <area shape="rect" coords="382,93,447,113" href="list.asp">
  <area shape="rect" coords="465,95,531,113" href="register.asp">
  <area shape="rect" coords="553,94,620,113" href="sendphoto.asp">
  <area shape="rect" coords="641,94,704,113" href="pics.asp">
</map>
</body>

</html>
