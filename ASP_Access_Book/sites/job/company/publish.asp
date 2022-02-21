<% Response.Buffer=True %>
<!--#include file="../inc/company.inc"-->
<!--#include file="../inc/html.inc"-->
<% uname=session("cuid") 
   modify=request("modify")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from company where uname='"&uname&"'and cname<>'""'"
rs.open sql,conn,1,1 %>

<% if rs.eof or rs.bof then
   response.write"<SCRIPT language=JavaScript>alert('您尚未登录公司资料，请先登录公司资料！');"
   response.write"javascript:history.go(-1);</SCRIPT>" 
   end if %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/register.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>天天人才—&gt;发布招聘信息</title>
</head>
<SCRIPT LANGUAGE="JavaScript">
<!--//
function check()
{
  if (register.job.value=="") {
		alert("请选择招聘职位名称！");
		register.job.focus();		
		    return (false);
  }
  if (register.zpnum.value=="") {
		alert("请输入招聘人数！");
		register.zpnum.focus();		
		    return (false);
  }
  if (isNaN(register.zpnum.value)){
		alert("请正确填写招聘人数！");
		register.zpnum.focus();		
		    return (false);
  }
  if (register.gzdd.value=="") {
		alert("请选择工作地点！");
		register.gzdd.focus();		
		    return (false);
  }
   if (register.zptext.value=="") {
		alert("请输入岗位描述！");
		register.zptext.focus();		
		    return (false);
  }
   if (register.xgyq.value=="") {
		alert("请输入相关要求！");
		register.xgyq.focus();		
		    return (false);
  }
else
	document.register.submit();
}
//-->
</SCRIPT>
<FORM name=register action=publish.asp method=post>
<body topmargin="0" leftmargin="0">
<!--#include file="../inc/top2.htm"--> 
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="87">
    <tr>
      <td width="780" height="18" valign="top" colspan="5" bgcolor="#53BEB0"> 
        　 </td>
    </tr>
    <tr>
      <td width="139" height="162" valign="top" bgcolor="#53BEB0"> 　 
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
      <td width="37" height="351" valign="top"><img border="0" src="../images/selfk.GIF"></td>
  </center>
      <td width="484" height="351">
        　
        <div align="left">
        <table border="1" cellpadding="0" cellspacing="0" width="92%" height="312" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
          <tr>
            <td width="100%" height="18" valign="bottom" bgcolor="#C6CEDE" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
              <p align="center">=== 发布招聘信息 ===</td>                               
          </tr>
          <tr>
            <td width="100%" height="305" valign="top" bgcolor="#F3F3F3">
              <p align="center">　<br>
              招聘职位:<select size="1" name="job"><OPTION>请在以下列表中选择</OPTION> 
			  <OPTION>-----------经营/管理类-----------
			  <OPTION value=(正/副)总裁/总经理/CEO <%if rs("job") ="(正/副)总裁/总经理/CEO" then Response.Write "selected"%>>(正/副)总裁/总经理/CEO
			  <OPTION value=企业策划 <%if rs("job") ="企业策划" then Response.Write "selected"%>>企业策划
			  <OPTION value=投资管理 <%if rs("job") ="投资管理" then Response.Write "selected"%>>投资管理
			  <OPTION value=企管部经理/企业管理顾问 <%if rs("job") ="企管部经理/企业管理顾问" then Response.Write "selected"%>>企管部经理/企业管理顾问
			  <OPTION value=企业管理人员 <%if rs("job") ="企业管理人员" then Response.Write "selected"%>>企业管理人员
	          <OPTION value=质管部经理/质量管理顾问 <%if rs("job") ="质管部经理/质量管理顾问" then Response.Write "selected"%>>质管部经理/质量管理顾问
			  <OPTION value=质管人员/质量工程师 <%if rs("job") ="质管人员/质量工程师" then Response.Write "selected"%>>质管人员/质量工程师
			  <OPTION value=技术经理/项目经理/CTO <%if rs("job") ="技术经理/项目经理/CTO" then Response.Write "selected"%>>技术经理/项目经理/CTO
			  <OPTION value=信息主管/CIO <%if rs("job") ="信息主管/CIO" then Response.Write "selected"%>>信息主管/CIO
			  <OPTION>--------------财务类--------------
			  <OPTION value=财务总监/经理/副经理/主任 <%if rs("job") ="财务总监/经理/副经理/主任" then Response.Write "selected"%>>财务总监/经理/副经理/主任
			  <OPTION value=会计师 <%if rs("job") ="会计师" then Response.Write "selected"%>>会计师<OPTION value=会计>会计
			  <OPTION value=出纳/收银 <%if rs("job") ="出纳/收银" then Response.Write "selected"%>>出纳/收银<OPTION value=统计 <%if rs("job") ="统计" then Response.Write "selected"%>>统计
			  <OPTION value=审计 <%if rs("job") ="审计" then Response.Write "selected"%>>审计
			  <OPTION>-----------销售/业务类-----------
			  <OPTION value=销售经理/副经理/主任 <%if rs("job") ="销售经理/副经理/主任" then Response.Write "selected"%>>销售经理/副经理/主任
			  <OPTION value=商务/贸易/国际业务 <%if rs("job") ="商务/贸易/国际业务" then Response.Write "selected"%>>商务/贸易/国际业务
			  <OPTION value=销售工程师 <%if rs("job") ="销售工程师" then Response.Write "selected"%>>销售工程师
			  <OPTION value=业务员/业务代表 <%if rs("job") ="业务员/业务代表" then Response.Write "selected"%>>业务员/业务代表
			  <OPTION value=报关 <%if rs("job") ="报关" then Response.Write "selected"%>>报关
			  <OPTION>--------------市场类--------------
			  <OPTION value=市场经理/副经理 <%if rs("job") ="市场经理/副经理" then Response.Write "selected"%>>市场经理/副经理
			  <OPTION value=市场/行销策划 <%if rs("job") ="市场/行销策划" then Response.Write "selected"%>>市场/行销策划
			  <OPTION value=市场调研/业务分析 <%if rs("job") ="市场调研/业务分析" then Response.Write "selected"%>>市场调研/业务分析
			  <OPTION value=公关、促销、礼仪 <%if rs("job") ="公关、促销、礼仪" then Response.Write "selected"%>>公关、促销、礼仪
			  <OPTION>--------------设计类--------------
			  <OPTION value=美术/图形设计 <%if rs("job") ="美术/图形设计" then Response.Write "selected"%>>美术/图形设计
			  <OPTION value=广告设计 <%if rs("job") ="广告设计" then Response.Write "selected"%>>广告设计
			  <OPTION value=文案策划 <%if rs("job") ="文案策划" then Response.Write "selected"%>>文案策划
			  <OPTION value=工业设计/产品设计 <%if rs("job") ="工业设计/产品设计" then Response.Write "selected"%>>工业设计/产品设计
			  <OPTION value=多媒体设计与制作 <%if rs("job") ="多媒体设计与制作" then Response.Write "selected"%>>多媒体设计与制作
			  <OPTION value=装潢设计 <%if rs("job") ="装潢设计" then Response.Write "selected"%>>装潢设计
			  <OPTION value=工艺品设计 <%if rs("job") ="工艺品设计" then Response.Write "selected"%>>工艺品设计
			  <OPTION value=纺织服装设计 <%if rs("job") ="纺织服装设计" then Response.Write "selected"%>>纺织服装设计
			  <OPTION value=家具/珠宝设计 <%if rs("job") ="家具/珠宝设计" then Response.Write "selected"%>>家具/珠宝设计
			  <OPTION value=电脑绘图人员 <%if rs("job") ="电脑绘图人员" then Response.Write "selected"%>>电脑绘图人员
			  <OPTION>------------客户服务类------------
			  <OPTION value=客服部经理/副经理 <%if rs("job") ="客服部经理/副经理" then Response.Write "selected"%>>客服部经理/副经理
			  <OPTION value=技术支持/客户培训 <%if rs("job") ="技术支持/客户培训" then Response.Write "selected"%>>技术支持/客户培训
			  <OPTION value=客服/热线咨询 <%if rs("job") ="客服/热线咨询" then Response.Write "selected"%>>客服/热线咨询
			  <OPTION value=前台/接待 <%if rs("job") ="前台/接待" then Response.Write "selected"%>>前台/接待
			  <OPTION>-----------行政/人事类-----------
			  <OPTION value=行政/人力资源经理 <%if rs("job") ="行政/人力资源经理" then Response.Write "selected"%>>行政/人力资源经理
			  <OPTION value=行政/人事人员 <%if rs("job") ="行政/人事人员" then Response.Write "selected"%>>行政/人事人员
			  <OPTION value=员工培训人员 <%if rs("job") ="员工培训人员" then Response.Write "selected"%>>员工培训人员
			  <OPTION value=企业文化/工会 <%if rs("job") ="企业文化/工会" then Response.Write "selected"%>>企业文化/工会
			  <OPTION>--------------文职类--------------
			  <OPTION value=英语翻译 <%if rs("job") ="英语翻译" then Response.Write "selected"%>>英语翻译
			  <OPTION value=其它外语翻译 <%if rs("job") ="其它外语翻译" then Response.Write "selected"%>>其它外语翻译
			  <OPTION value=图书情报/资料管理 <%if rs("job") ="图书情报/资料管理" then Response.Write "selected"%>>图书情报/资料管理
			  <OPTION value=技术资料编写 <%if rs("job") ="技术资料编写" then Response.Write "selected"%>>技术资料编写
			  <OPTION value=文秘/高级文员 <%if rs("job") ="文秘/高级文员" then Response.Write "selected"%>>文秘/高级文员
			  <OPTION value=文员/电脑打字员/操作员 <%if rs("job") ="文员/电脑打字员/操作员" then Response.Write "selected"%>>文员/电脑打字员/操作员
			  <OPTION>-----------工业/工厂类-----------
			  <OPTION value=厂长/副厂长 <%if rs("job") ="厂长/副厂长" then Response.Write "selected"%>>厂长/副厂长
			  <OPTION value=生产管理 <%if rs("job") ="生产管理" then Response.Write "selected"%>>生产管理
			  <OPTION value=工程管理 <%if rs("job") ="工程管理" then Response.Write "selected"%>>工程管理
			  <OPTION value=品质管理 <%if rs("job") ="品质管理" then Response.Write "selected"%>>品质管理
			  <OPTION value=物料管理 <%if rs("job") ="物料管理" then Response.Write "selected"%>>物料管理
			  <OPTION value=设备管理 <%if rs("job") ="设备管理" then Response.Write "selected"%>>设备管理
			  <OPTION value=采购管理 <%if rs("job") ="采购管理" then Response.Write "selected"%>>采购管理
			  <OPTION value=仓库管理 <%if rs("job") ="库管理" then Response.Write "selected"%>>仓库管理
			  <OPTION value=计划员 <%if rs("job") ="计划员" then Response.Write "selected"%>>计划员
			  <OPTION value=化验工作 <%if rs("job") ="化验工作" then Response.Write "selected"%>>化验工作
			  <OPTION value=技工 <%if rs("job") ="技工" then Response.Write "selected"%>>技工
			  <OPTION value=普工 <%if rs("job") ="普工" then Response.Write "selected"%>>普工
			  <OPTION>-----------后勤/服务类-----------
			  <OPTION value=司机 <%if rs("job") ="司机" then Response.Write "selected"%>>司机
			  <OPTION value=保安/厨师/清洁工 <%if rs("job") ="保安/厨师/清洁工 " then Response.Write "selected"%>>保安/厨师/清洁工
			  <OPTION value=服务员 <%if rs("job") ="服务员 " then Response.Write "selected"%>>服务员
			  <OPTION value=营业员 <%if rs("job") ="营业员 " then Response.Write "selected"%>>营业员
			  <OPTION>----------计算机专业人员----------
			  <OPTION value=系统分析员 <%if rs("job") ="系统分析员" then Response.Write "selected"%>>系统分析员
			  <OPTION value=软件（测试）工程师 <%if rs("job") ="软件（测试）工程师 " then Response.Write "selected"%>>软件（测试）工程师
			  <OPTION value=硬件（测试）工程师 <%if rs("job") ="硬件（测试）工程师" then Response.Write "selected"%>>硬件（测试）工程师
			  <OPTION value=系统工程师/网管 <%if rs("job") ="系统工程师/网管" then Response.Write "selected"%>>系统工程师/网管
			  <OPTION value=网站策划/信息编辑 <%if rs("job") ="网站策划/信息编辑" then Response.Write "selected"%>>网站策划/信息编辑
			  <OPTION value=网页设计/电脑美工 <%if rs("job") ="网页设计/电脑美工" then Response.Write "selected"%>>网页设计/电脑美工
			  <OPTION value=Internet/WEB/电子商务开发 <%if rs("job") ="Internet/WEB/电子商务开发" then Response.Write "selected"%>>Internet/WEB/电子商务开发
			  <OPTION>-------电子/通讯类专业人员-------
			  <OPTION value=电子工程师 <%if rs("job") ="电子工程师" then Response.Write "selected"%>>电子工程师
			  <OPTION value=电子元器件工程师 <%if rs("job") ="电子元器件工程师" then Response.Write "selected"%>>电子元器件工程师
			  <OPTION value=自动控制 <%if rs("job") ="自动控制" then Response.Write "selected"%>>自动控制
			  <OPTION value=智能大厦/综合布线/弱电 <%if rs("job") ="智能大厦/综合布线/弱电" then Response.Write "selected"%>>智能大厦/综合布线/弱电
			  <OPTION value=仪器仪表/计量 <%if rs("job") ="仪器仪表/计量" then Response.Write "selected"%>>仪器仪表/计量
			  <OPTION value=电气 <%if rs("job") ="电气" then Response.Write "selected"%>>电气
			  <OPTION value=电力 <%if rs("job") ="电力" then Response.Write "selected"%>>电力
			  <OPTION value=通讯工程师 <%if rs("job") ="通讯工程师" then Response.Write "selected"%>>通讯工程师
			  <OPTION value=单片机/DSP/底层软件开发 <%if rs("job") ="单片机/DSP/底层软件开发" then Response.Write "selected"%>>单片机/DSP/底层软件开发
			  <OPTION>-----------机械专业人员-----------
			  <OPTION value=机械工程师 <%if rs("job") ="机械工程师" then Response.Write "selected"%>>机械工程师
			  <OPTION value=模具工程师 <%if rs("job") ="模具工程师" then Response.Write "selected"%>>模具工程师
			  <OPTION value=机电工程师 <%if rs("job") ="机电工程师" then Response.Write "selected"%>>机电工程师
			  <OPTION value=各种车辆/飞行器设计 <%if rs("job") ="各种车辆/飞行器设计" then Response.Write "selected"%>>各种车辆/飞行器设计
			  <OPTION>-------房地产/建筑专业人员-------
			  <OPTION value=房地产开发/策划 <%if rs("job") ="房地产开发/策划" then Response.Write "selected"%>>房地产开发/策划
			  <OPTION value=房地产评估/交易 <%if rs("job") ="房地产评估/交易" then Response.Write "selected"%>>房地产评估/交易
			  <OPTION value=建筑/结构工程师 <%if rs("job") ="建筑/结构工程师" then Response.Write "selected"%>>建筑/结构工程师
			  <OPTION value=工程监理师 <%if rs("job") ="工程监理师" then Response.Write "selected"%>>工程监理师
			  <OPTION value=工程预决算 <%if rs("job") ="工程预决算" then Response.Write "selected"%>>工程预决算
			  <OPTION value=给排水/水电工程师 <%if rs("job") ="给排水/水电工程师" then Response.Write "selected"%>>给排水/水电工程师
			  <OPTION value=制冷暖通 <%if rs("job") ="制冷暖通" then Response.Write "selected"%>>制冷暖通
			  <OPTION value=物业管理 <%if rs("job") ="物业管理" then Response.Write "selected"%>>物业管理
			  <OPTION>--------金融/经济专业人员--------
			  <OPTION value=金融业 <%if rs("job") ="金融业" then Response.Write "selected"%>>金融业
			  <OPTION value=证券期货 <%if rs("job") ="证券期货" then Response.Write "selected"%>>证券期货
			  <OPTION value=保险业 <%if rs("job") ="保险业" then Response.Write "selected"%>>保险业
			  <OPTION value=税务人员 <%if rs("job") ="税务人员" then Response.Write "selected"%>>税务人员
			  <OPTION value=其它金融/经济人员 <%if rs("job") ="其它金融/经济人员" then Response.Write "selected"%>>其它金融/经济人员
			  <OPTION>------文教体卫/法律专业人员------
			  <OPTION value=新闻/出版 <%if rs("job") ="新闻/出版" then Response.Write "selected"%>>新闻/出版
			  <OPTION value=广播电视/文化艺术 <%if rs("job") ="广播电视/文化艺术" then Response.Write "selected"%>>广播电视/文化艺术
			  <OPTION value=高等教育 <%if rs("job") ="高等教育" then Response.Write "selected"%>>高等教育
			  <OPTION value=中等教育 <%if rs("job") ="中等教育" then Response.Write "selected"%>>中等教育
			  <OPTION value=小学/幼儿教育 <%if rs("job") ="小学/幼儿教育" then Response.Write "selected"%>>小学/幼儿教育
			  <OPTION value=职业教育/培训/家教 <%if rs("job") ="职业教育/培训/家教" then Response.Write "selected"%>>职业教育/培训/家教
			  <OPTION value=体育类 <%if rs("job") ="体育类" then Response.Write "selected"%>>体育类
			  <OPTION value=卫生医疗 <%if rs("job") ="卫生医疗" then Response.Write "selected"%>>卫生医疗
			  <OPTION value=律师/法律顾问 <%if rs("job") ="律师/法律顾问" then Response.Write "selected"%>>律师/法律顾问
			  <OPTION>----------服务业专业人员----------
			  <OPTION value=旅游/导游 <%if rs("job") ="旅游/导游" then Response.Write "selected"%>>旅游/导游
			  <OPTION value=酒店/餐饮 <%if rs("job") ="酒店/餐饮" then Response.Write "selected"%>>酒店/餐饮
			  <OPTION value=寻呼/声讯 <%if rs("job") ="寻呼/声讯" then Response.Write "selected"%>>寻呼/声讯
			  <OPTION value=其它服务行业 <%if rs("job") ="其它服务行业" then Response.Write "selected"%>>其它服务行业
			  <OPTION>----------其它专业人员 ----------
			  <OPTION value=动力/能源 <%if rs("job") ="动力/能源" then Response.Write "selected"%>>动力/能源
			  <OPTION value=声光学技术 <%if rs("job") ="声光学技术" then Response.Write "selected"%>>声光学技术
			  <OPTION value=化工技术 <%if rs("job") ="化工技术" then Response.Write "selected"%>>化工技术
			  <OPTION value=生物制药 <%if rs("job") ="生物制药" then Response.Write "selected"%>>生物制药
			  <OPTION value=测绘技术 <%if rs("job") ="测绘技术" then Response.Write "selected"%>>测绘技术
			  <OPTION value=道桥技术 <%if rs("job") ="" then Response.Write "selected"%>>道桥技术
			  <OPTION value=环境/城市规划 <%if rs("job") ="环境/城市规划" then Response.Write "selected"%>>环境/城市规划
			  <OPTION value=地质/矿产 <%if rs("job") ="地质/矿产" then Response.Write "selected"%>>地质/矿产
			  <OPTION value=粮食/食品/糖酒 <%if rs("job") ="粮食/食品/糖酒" then Response.Write "selected"%>>粮食/食品/糖酒
			  <OPTION value=纺织服装技术 <%if rs("job") ="纺织服装技术" then Response.Write "selected"%>>纺织服装技术
			  <OPTION value=包装/印刷/造纸 <%if rs("job") ="包装/印刷/造纸" then Response.Write "selected"%>>包装/印刷/造纸
			  <OPTION value=冶金/喷涂/金属材料 <%if rs("job") ="冶金/喷涂/金属材料" then Response.Write "selected"%>>冶金/喷涂/金属材料
			  <OPTION value=安全消防 <%if rs("job") ="安全消防" then Response.Write "selected"%>>安全消防
			  <OPTION value=农林渔牧/园林/园艺 <%if rs("job") ="农林渔牧/园林/园艺" then Response.Write "selected"%>>农林渔牧/园林/园艺
			  <OPTION value=交通运输（海陆空） <%if rs("job") ="交通运输（海陆空）" then Response.Write "selected"%>>交通运输（海陆空）</select></p>
              <p align="center">招聘人数:<INPUT maxLength=3    
                  size=3 name=zpnum value="<%=rs("zpnum")%>">人 工作地点:<B><SELECT size=1 name=gzdd>               
              <OPTION>请在以下列表中选择</OPTION>
			  <OPTION value=北京市 <%if rs("gzdd") ="北京市" then Response.Write "selected"%>>北京市
              <OPTION value=天津市 <%if rs("gzdd") ="天津市" then Response.Write "selected"%>>天津市
              <OPTION value=上海市 <%if rs("gzdd") ="上海市" then Response.Write "selected"%>>上海市
			  <OPTION value=重庆市 <%if rs("gzdd") ="重庆市" then Response.Write "selected"%>>重庆市
              <OPTION value=安徽省 <%if rs("gzdd") ="安徽省" then Response.Write "selected"%>>安徽省
              <OPTION value=福建省 <%if rs("gzdd") ="福建省" then Response.Write "selected"%>>福建省
              <OPTION value=甘肃省 <%if rs("gzdd") ="甘肃省" then Response.Write "selected"%>>甘肃省
              <OPTION value=广东省 <%if rs("gzdd") ="广东省" then Response.Write "selected"%>>广东省
			  <OPTION value=四川省 <%if rs("gzdd") ="四川省" then Response.Write "selected"%>>四川省
              <OPTION value=贵州省 <%if rs("gzdd") ="贵州省" then Response.Write "selected"%>>贵州省
              <OPTION value=湖北省 <%if rs("gzdd") ="湖北省" then Response.Write "selected"%>>湖北省
              <OPTION value=湖南省 <%if rs("gzdd") ="湖南省" then Response.Write "selected"%>>湖南省
			  <OPTION value=吉林省 <%if rs("gzdd") ="吉林省" then Response.Write "selected"%>>吉林省
              <OPTION value=江苏省 <%if rs("gzdd") ="江苏省" then Response.Write "selected"%>>江苏省
              <OPTION value=江西省 <%if rs("gzdd") ="江西省" then Response.Write "selected"%>>江西省
              <OPTION value=辽宁省 <%if rs("gzdd") ="辽宁省" then Response.Write "selected"%>>辽宁省
              <OPTION value=青海省 <%if rs("gzdd") ="青海省" then Response.Write "selected"%>>青海省
			  <OPTION value=山东省 <%if rs("gzdd") ="山东省" then Response.Write "selected"%>>山东省
              <OPTION value=山西省 <%if rs("gzdd") ="山西省" then Response.Write "selected"%>>山西省
			  <OPTION value=陕西省 <%if rs("gzdd") ="陕西省" then Response.Write "selected"%>>陕西省
              <OPTION value=云南省 <%if rs("gzdd") ="云南省" then Response.Write "selected"%>>云南省
			  <OPTION value=浙江省 <%if rs("gzdd") ="浙江省" then Response.Write "selected"%>>浙江省
			  <OPTION value=黑龙江省 <%if rs("gzdd") ="黑龙江省" then Response.Write "selected"%>>黑龙江省
			
             </SELECT></B>
              </p>
              <p align="center">岗位描述:<br>
             <%if rs("zptext")<>"" then
			 kzptext=replace(rs("zptext"),"<br>",chr(13))
             kzptext=replace(kzptext,"&nbsp;"," ")
			 else
			 kzptext="" %>
             <%end if%>      
			 <textarea rows="6" name="zptext" cols="41"><%=kzptext%></textarea>      
             </p>
             <p align="center">相关要求:<br>
             <%if rs("xgyq")<>"" then
	         kxgyq=replace(rs("xgyq"),"<br>",chr(13))
             kxgyq=replace(kxgyq,"&nbsp;"," ")
			 else
			 kxgyq=""%>
             <%end if%>      
			 <textarea rows="6" name="xgyq" cols="41"><%=kxgyq%></textarea>      
              </p>
              <p align="center"><input type="button" value="发 布" onClick="check()">

              <p align="center">　</td>
          </tr>
        </table>
        </div>
        <p>　
      </td>
  <center>
      <td width="1" height="351" valign="top" bgcolor="#00006A"></td>
      <td width="124" height="351" valign="top" bgcolor="#F3F3F3">　</td>
    </tr>
    <tr>
      <td width="780" height="1" valign="top" colspan="5" bgcolor="#53BEB0"> 
        <p align="center"> </td>
    </tr>
    <tr>
      <td width="780" height="3" valign="top" colspan="5">
      <p align="center">
      </td>
    </tr>
    <tr>
      <td width="780" height="7" valign="top" colspan="5">
      <p align="center"><script language="javascript" src="../inc/copyright.js"></script>
      </td>
    </tr>
    <tr>
      <td width="780" height="3" valign="top" colspan="5">
      </td>
    </tr>
  </table>
  </center>
</div>
</form>
</body>

</html>

<% zptext=htmlencode2(request("zptext"))
if zptext="" then Response.End
   job=request("job")
   zpnum=request("zpnum")
   gzdd=request("gzdd")
   xgyq=htmlencode2(request("xgyq"))

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from company where uname='"&uname&"'"
rs.open sql,conn,3,3
rs("zptext")=zptext
rs("job")=job
rs("zpnum")=zpnum
rs("gzdd")=gzdd
rs("xgyq")=xgyq
rs("idate")=date()
rs.update
rs.close
Response.Redirect "main.asp"
%>
