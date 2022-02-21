<% Response.Buffer=True %>
<!--#include file="../inc/company.inc"-->
<!--#include file="../inc/html.inc"-->
<% uname=session("cuid")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from company where uname='"&uname&"'"
rs.open sql,conn,1,1 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/register.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>天天人才—&gt;等录公司资料</title>
</head>
<SCRIPT language=JavaScript src="../inc/validate.js"></SCRIPT>
<SCRIPT language=JavaScript src="../inc/companyreg.js"></SCRIPT>
<FORM name=register action=register.asp method=post>
<body topmargin="0" leftmargin="0">
<!--#include file="../inc/top2.htm"--> 
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="61">
    <tr>
      <td width="778" height="18" valign="top" colspan="5" bgcolor="#53BEB0"> 
        　 </td>
    </tr>
    <tr>
      <td width="139" valign="top" bgcolor="#53BEB0"> 　 
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="118" height="280">
            <tr>
              <td width="100%" height="163" background="../images/stat-bg.GIF" valign="top">
        <p align="center"><br>
        <a href="main.asp">登录首页</a><br>
        <br>
		<a href="register.asp">登录/更新公司资料</a><br>
        <br>
		<a href="publish.asp">发布/更新招聘信息</a><br>
        <br>
        <a href="../changepwd.asp?stype=company" target="_blank">修改登录密码</a><br>
        <br>
        <a href="search.asp">全部人才列表</a>
              </td>
            </tr>
            <tr>
              <td width="100%" height="117" background="../images/stat-bg.GIF" valign="top">
        <p align="center"><br>
        <br>
        <a href="favorite.asp">我的收藏夹<br>
        <br>
        <a href="mailbox.asp">我的信箱</a><br>
        <br>
        <a href="../exit.asp">退出登录</a>
              </td>
            </tr>
          </table>
        </div>
        <p align="center"> </td>
      <td width="31" height="325" valign="top"><img border="0" src="../images/selfk.GIF"></td>
  </center>
      <td width="481" height="325">
        　
        <div align="left">
        <table border="1" cellpadding="0" cellspacing="0" width="94%" height="263" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
          <tr>
            <td width="100%" height="18" valign="bottom" bgcolor="#C6CEDE" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
              <p align="center">=== 公司详细资料 ===</td>                           
          </tr>
          <tr>
            <td width="100%" height="20" valign="bottom">
              &nbsp;用户名:<FONT COLOR="#FF0000"><%=uname%></FONT> &nbsp;</td>                           
          </tr>
          <tr>
            
			<td width="100%" height="22" valign="bottom" bgcolor="#F3F3F3">
            &nbsp;公司名称:<%if rs("cname")<>"" then%>
			<%=rs("cname")%><%else%>    
			<input type="text" name="cname" size="40" maxLength=40 value="<%=rs("cname")%>">    
			<%end if%>    
			</td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom">
              <p align="left">&nbsp;所属行业:<SELECT size=1 name=trade> 
              <OPTION>请在以下列表中选择</OPTION> <OPTION 
              value=计算机业 <%if rs("trade") ="计算机业" then Response.Write "selected"%>>计算机业</OPTION> <OPTION value=电子电器、通讯设备 <%if rs("trade") ="电子电器、通讯设备" then Response.Write "selected"%>>电子电器、通讯设备</OPTION> 
              <OPTION value=科研设计、科技开发  <%if rs("trade") ="科研设计、科技开发 " then Response.Write "selected"%>>科研设计、科技开发</OPTION> <OPTION 
              value=机电设备、电力 <%if rs("trade") ="机电设备、电力 " then Response.Write "selected"%>>机电设备、电力</OPTION> <OPTION value=仪器仪表、度量衡器 <%if rs("trade") ="仪器仪表、度量衡器" then Response.Write "selected"%>>仪器仪表、度量衡器</OPTION> 
              <OPTION value=机械制造及设备 <%if rs("trade") ="机械制造及设备" then Response.Write "selected"%>>机械制造及设备</OPTION> <OPTION 
              value=各种车辆制造与销售 <%if rs("trade") ="各种车辆制造与销售" then Response.Write "selected"%>>各种车辆制造与销售</OPTION> <OPTION value=动力、电力 <%if rs("trade") ="动力、电力" then Response.Write "selected"%>>动力、电力</OPTION> 
              <OPTION value=化学化工、生物制药 <%if rs("trade") ="化学化工、生物制药" then Response.Write "selected"%>>化学化工、生物制药</OPTION> <OPTION 
              value=五金矿产、金属制品 <%if rs("trade") ="五金矿产、金属制品" then Response.Write "selected"%>>五金矿产、金属制品</OPTION> <OPTION value=纺织、服装 <%if rs("trade") ="纺织、服装" then Response.Write "selected"%>>纺织、服装</OPTION> 
              <OPTION value=农林牧副渔 <%if rs("trade") ="农林牧副渔" then Response.Write "selected"%>>农林牧副渔</OPTION> <OPTION value=轻工 <%if rs("trade") ="轻工" then Response.Write "selected"%>>轻工</OPTION> 
              <OPTION value=房地产、建筑、装潢 <%if rs("trade") ="房地产、建筑、装潢" then Response.Write "selected"%>>房地产、建筑、装潢</OPTION> <OPTION 
              value=造纸、印刷、包装 <%if rs("trade") ="造纸、印刷、包装" then Response.Write "selected"%>>造纸、印刷、包装</OPTION> <OPTION value=木材、家具 <%if rs("trade") ="木材、家具" then Response.Write "selected"%>>木材、家具</OPTION> 
              <OPTION value=广告、策划、设计 <%if rs("trade") ="广告、策划、设计" then Response.Write "selected"%>>广告、策划、设计</OPTION> <OPTION 
              value=新闻出版、广播电视 <%if rs("trade") ="新闻出版、广播电视" then Response.Write "selected"%>>新闻出版、广播电视</OPTION> <OPTION 
              value=信息咨询、人才交流 <%if rs("trade") ="信息咨询、人才交流" then Response.Write "selected"%>>信息咨询、人才交流</OPTION> <OPTION 
              value=金融、保险 <%if rs("trade") ="金融、保险" then Response.Write "selected"%>>金融、保险</OPTION> <OPTION value=交通运输 <%if rs("trade") ="交通运输" then Response.Write "selected"%>>交通运输</OPTION> <OPTION 
              value=邮政、电信 <%if rs("trade") ="邮政、电信" then Response.Write "selected"%>>邮政、电信</OPTION> <OPTION value=综合性工商企业 <%if rs("trade") ="综合性工商企业" then Response.Write "selected"%>>综合性工商企业</OPTION> 
              <OPTION value=市政、公用事业 <%if rs("trade") ="市政、公用事业" then Response.Write "selected"%>>市政、公用事业</OPTION> <OPTION 
              value=教育事业 <%if rs("trade") ="教育事业" then Response.Write "selected"%>>教育事业</OPTION> <OPTION value=医疗、卫生事业 <%if rs("trade") ="医疗、卫生事业" then Response.Write "selected"%>>医疗、卫生事业</OPTION> 
              <OPTION value=文化、艺术 <%if rs("trade") ="文化、艺术" then Response.Write "selected"%>>文化、艺术</OPTION> <OPTION value=文娱体育、办公用品 <%if rs("trade") ="文娱体育、办公用品" then Response.Write "selected"%>>文娱体育、办公用品</OPTION> <OPTION value=日用生活服务 <%if rs("trade") ="日用生活服务" then Response.Write "selected"%>>日用生活服务</OPTION> 
              <OPTION value=商业服务 <%if rs("trade") ="商业服务" then Response.Write "selected"%>>商业服务</OPTION> <OPTION value=物资供销 <%if rs("trade") ="物资供销" then Response.Write "selected"%>>物资供销</OPTION> 
              <OPTION value=饮食、旅游业 <%if rs("trade") ="饮食、旅游业" then Response.Write "selected"%>>饮食、旅游业</OPTION> <OPTION 
              value=粮油、副食 <%if rs("trade") ="粮油、副食" then Response.Write "selected"%>>粮油、副食</OPTION></SELECT>
              </p>
            </td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom" bgcolor="#F3F3F3">
              &nbsp;企业性质:<SELECT size=1 name=cxz> <OPTION 
              value="" selected>请在以下列表中选择</OPTION> <OPTION 
              value=行政机关 <%if rs("cxz") ="行政机关" then Response.Write "selected"%>>行政机关</OPTION> <OPTION value=社会团体 <%if rs("cxz") ="社会团体" then Response.Write "selected"%>>社会团体</OPTION> <OPTION 
              value=事业单位 <%if rs("cxz") ="事业单位" then Response.Write "selected"%>>事业单位</OPTION> <OPTION value=国有企业 <%if rs("cxz") ="国有企业" then Response.Write "selected"%>>国有企业</OPTION> <OPTION 
              value=三资企业 <%if rs("cxz") ="三资企业" then Response.Write "selected"%>>三资企业</OPTION> <OPTION value=集体企业 <%if rs("cxz") ="集体企业" then Response.Write "selected"%>>集体企业</OPTION> <OPTION value=乡镇企业 <%if rs("cxz") ="乡镇企业" then Response.Write "selected"%>>乡镇企业</OPTION> <OPTION 
              value=私营企业 <%if rs("cxz") ="私营企业" then Response.Write "selected"%>>私营企业</OPTION> </SELECT>
            </td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom">
              &nbsp;所在区域:<B><SELECT size=1 name=area> 
              <OPTION>请在以下列表中选择</OPTION>
			  <OPTION value=北京市 <%if rs("area") ="北京市" then Response.Write "selected"%>>北京市
              <OPTION value=天津市 <%if rs("area") ="天津市" then Response.Write "selected"%>>天津市
              <OPTION value=上海市 <%if rs("area") ="上海市" then Response.Write "selected"%>>上海市
			  <OPTION value=重庆市 <%if rs("area") ="重庆市" then Response.Write "selected"%>>重庆市
              <OPTION value=安徽省 <%if rs("area") ="安徽省" then Response.Write "selected"%>>安徽省
              <OPTION value=福建省 <%if rs("area") ="福建省" then Response.Write "selected"%>>福建省
              <OPTION value=甘肃省 <%if rs("area") ="甘肃省" then Response.Write "selected"%>>甘肃省
              <OPTION value=广东省 <%if rs("area") ="广东省" then Response.Write "selected"%>>广东省
			  <OPTION value=四川省 <%if rs("area") ="四川省" then Response.Write "selected"%>>四川省
              <OPTION value=贵州省 <%if rs("area") ="贵州省" then Response.Write "selected"%>>贵州省
              <OPTION value=湖北省 <%if rs("area") ="湖北省" then Response.Write "selected"%>>湖北省
              <OPTION value=湖南省 <%if rs("area") ="湖南省" then Response.Write "selected"%>>湖南省
			  <OPTION value=吉林省 <%if rs("area") ="吉林省" then Response.Write "selected"%>>吉林省
              <OPTION value=江苏省 <%if rs("area") ="江苏省" then Response.Write "selected"%>>江苏省
              <OPTION value=江西省 <%if rs("area") ="江西省" then Response.Write "selected"%>>江西省
              <OPTION value=辽宁省 <%if rs("area") ="辽宁省" then Response.Write "selected"%>>辽宁省
              <OPTION value=青海省 <%if rs("area") ="青海省" then Response.Write "selected"%>>青海省
			  <OPTION value=山东省 <%if rs("area") ="山东省" then Response.Write "selected"%>>山东省
              <OPTION value=山西省 <%if rs("area") ="山西省" then Response.Write "selected"%>>山西省
			  <OPTION value=陕西省 <%if rs("area") ="陕西省" then Response.Write "selected"%>>陕西省
              <OPTION value=云南省 <%if rs("area") ="云南省" then Response.Write "selected"%>>云南省
			  <OPTION value=浙江省 <%if rs("area") ="浙江省" then Response.Write "selected"%>>浙江省
			  <OPTION value=黑龙江省 <%if rs("area") ="黑龙江省" then Response.Write "selected"%>>黑龙江省

              </SELECT></B>
            </td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom" bgcolor="#F3F3F3">
              &nbsp;成立日期:<input type="text" name="fdate" size="10" maxLength=10 value="<%=rs("fdate")%>">
            </td>                 
          </tr>
          <tr>
            <% kfund=rs("fund")
			  if kfund="未知" then kfund="" end if %>
            <td width="100%" height="28" valign="bottom">
              &nbsp;注册资金;<input type="text" name="fund" size="6" maxLength=6 value="<%=kfund%>">万人民币</td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom" bgcolor="#F3F3F3">
              &nbsp;通讯地址:<input type="text" name="address" size="30" maxLength=40 value="<%=rs("address")%>">
            </td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom">
              &nbsp;邮政编码;<input type="text" name="zip" size="6" size="6" maxLength=6 value="<%=rs("zip")%>">
            </td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom" bgcolor="#F3F3F3">
              <p align="left">&nbsp;联系人:&nbsp; <input type="text" name="pname" size="20" maxLength=20 value="<%=rs("pname")%>">    
              </p></td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom">
              &nbsp;联系电话:<input type="text" name="phone" size="20" maxLength=40 value="<%=rs("phone")%>">
            </td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom" bgcolor="#F3F3F3">
              <p align="left">&nbsp;传真号码:<input type="text" name="fax" size="20" maxLength=40 value="<%=rs("fax")%>">
              </p></td>                 
          </tr>
          <tr>
            <td width="100%" height="28" valign="bottom">
              &nbsp;电子信箱:<input type="text" name="email" size="20" maxLength=40 value="<%=rs("email")%>">
            </td>                 
          </tr>
          <tr>
          <td width="100%" height="28" valign="bottom" bgcolor="#F3F3F3">
           &nbsp;公司网站:<input type="text" name="http" size="20" maxLength=40 
		   <% if rs("http") <>"" then%> value="<%=rs("http")%>" <%else%> value="http://"<%end if%>>
            </td>                 
          </tr>
          <tr>
            <td width="100%" valign="top">
              <p align="center"><br>
              公司简介:<br>
              <br>
              <%if rs("cname")<>"" then 
			  kjianj=replace(rs("jianj"),"<br>",chr(13))
              kjianj=replace(kjianj,"&nbsp;"," ")
			  else
			  kjianj="" %>
			  <%end if%>      
   <textarea rows="9" name="jianj" cols="62" style="background-color: #F3F3F3; color: #000060"><%=kjianj%></textarea>    
              </p>
              <p align="center"><input type="button" value="确 定" <%if rs("cname")<>"" then%>onClick="checkmod()"<%else%>onClick="check()"<%end if%>>
              <br>
              <br>
              </p></td>
          </tr>
        </table>
        </div>
      </td>
  <center>
      <td width="1" height="325" valign="top" bgcolor="#00006A"></td>
      <td width="122" height="325" valign="top" bgcolor="#F3F3F3">　</td>
    </tr>
    <tr>
      <td width="778" height="1" valign="top" colspan="5" bgcolor="#53BEB0"> 
        <p align="center"> </td>
    </tr>
    <tr>
      <td width="778" height="3" valign="top" colspan="5">
      <p align="center">
      </td>
    </tr>
    <tr>
      <td width="778" height="7" valign="top" colspan="5">
      <p align="center"><script language="javascript" src="../inc/copyright.js"></script>
      </td>
    </tr>
    <tr>
      <td width="778" height="3" valign="top" colspan="5">
      </td>
    </tr>
  </table>
  </center>
</div>
</form>
</body>

</html>

<% pname=request("pname") 
if pname="" then Response.End
trade=request("trade")
cxz=request("cxz")
area=request("area")
fdate=request("fdate")
fund=request("fund")
cname=request("cname")
address=request("address")
zip=request("zip")
phone=request("phone")
fax=request("fax")
email=request("email")
http=request("http")
jianj=htmlencode2(request("jianj"))
if fund="" then fund="未知" end if
if cname="" then cname=rs("cname") end if
if fax="" then fax="未知" end if
if http="" then http="http://" end if
rs.close
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from company where uname='"&uname&"'"
rs.open sql,conn,3,3
rs("trade")=trade
rs("cxz")=cxz
rs("area")=area
rs("fdate")=fdate
rs("cname")=cname
rs("fund")=fund
rs("pname")=pname
rs("address")=address
rs("zip")=zip
rs("phone")=phone
rs("fax")=fax
rs("email")=email
rs("http")=http
rs("jianj")=jianj
rs.update
rs.close
Response.Redirect "main.asp" %>