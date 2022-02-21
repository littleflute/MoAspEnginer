<% Response.Buffer=True %>
<!--#include file="inc/dbconn.inc"-->
<% 
set rs=server.createobject("adodb.recordset")
sql="select * from company"
rs.open sql,conn,1,1
r1=rs.recordcount
rs.close                                                                    
set rs=nothing
set rs=server.createobject("adodb.recordset")
sql="select * from company where job<>'""' order by id desc"           
rs.open sql,conn,1,1 
jobnum=rs.recordcount %> 

<html>
<title>天天人才—&gt;首页</title>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" href="inc/index.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

</head>

<body topmargin="0" leftmargin="0">
<!--#include file="inc/top.htm"--> 
<SCRIPT language=JavaScript src="inc/f-index.js"></SCRIPT>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="120">
    <tr bgcolor="#53BEB0"> 
      <td width="778" height="18" colspan="5" valign="top"> 　 </td>
    </tr>
    <tr>
      <td height="46" valign="top" bgcolor="#53BEB0" rowspan="2" width="152"> 
        <div align="center">
          <center>
          <table border="0" cellpadding="0" cellspacing="0" width="140" height="142" background="images/loginbg.GIF">
            <tr>
              <td width="136" height="19" valign="top">　</td>
            </tr>
            <form name=login action=login.asp method=post>
			<tr>
              <td width="136" height="47" valign="top">&nbsp;&nbsp;&nbsp;&nbsp; 用户名:<input type="text" name="uname" size="7" maxLength=20 style="font-family: 宋体; font-size: 9pt; background-color: #EBEBEB; color: #00006A"><br>                                                                                                     
                &nbsp;&nbsp;&nbsp;&nbsp; 密&nbsp; 码:<input type="password" name="pwd" maxLength=20 size="7" style="font-family: 宋体; font-size: 9pt; background-color: #EBEBEB; color: #00006A"></td>                                                                                                      
            </tr>
            <tr>
              <td width="136" height="14" valign="top">&nbsp;&nbsp;&nbsp; <input type="radio" value="person" checked name="usertype">个人                                                                                                      
                <input type="radio" name="usertype" value="company">单位</td>
            </tr>
          </center>
  </center>
            <tr>
              <td width="136" height="71" valign="top">
                <p align="center"><br>
                <input type="button" value="登 录" name="B1" style="position: relative; color: #00006A; font-family: 宋体; font-size: 9pt; height: 19" onClick="check()"><br>
                <br>
                <a href="addnew.asp"><font color="#00006A">新用户注册</font></a></td>
            </tr>
          </table>
  <center>
          <center></form>
            </center>
        </div>
        <div align="center">
          <center>
          <form name=search action=search.asp method=post target="_blank">
		  <table border="0" cellpadding="0" cellspacing="0" width="138" height="168" background="images/search-bg.GIF">
            <tr>
                      <td width="136" height="167">
                        <p align="center"><font color="#0000FF"><br>
                        </font><font color="#000055">=== 站内搜索 ===</font><font color="#000000"><br>                              
                        <br>
                        </font><SELECT 
                  size=1 name=stype style="background-color: #EBEBEB; color: #00006A; font-family: 宋体; font-size: 9pt">
                  <OPTION value="" selected>请选择搜索类别</option>
				  <option value="company">职位搜索</option>
                  <option value="person">人才搜索</option>
                </SELECT>
                        <p align="center"><SELECT size=1 name=gzdd style="color: #00006A; font-family: 宋体; font-size: 9pt; background-color: #EBEBEB"> 
              <OPTION selected>请选择所在地区</OPTION>
			  <OPTION value=noxz>不&nbsp; 限
			  <OPTION value=北京市>北京市
              <OPTION value=天津市>天津市
              <OPTION value=上海市>上海市
			  <OPTION value=重庆市>重庆市
              <OPTION value=安徽省>安徽省
              <OPTION value=福建省>福建省
              <OPTION value=甘肃省>甘肃省
              <OPTION value=广东省>广东省
			  <OPTION value=四川省>四川省
              <OPTION value=贵州省>贵州省
              <OPTION value=湖北省>湖北省
              <OPTION value=湖南省>湖南省
			  <OPTION value=吉林省>吉林省
              <OPTION value=江苏省>江苏省
              <OPTION value=江西省>江西省
              <OPTION value=辽宁省>辽宁省
              <OPTION value=青海省>青海省
			  <OPTION value=山东省>山东省
              <OPTION value=山西省>山西省
			  <OPTION value=陕西省>陕西省
              <OPTION value=云南省>云南省
			  <OPTION value=浙江省>浙江省
			  <OPTION value=黑龙江省>黑龙江省
			  </SELECT></p>
            <p align="center"><INPUT 
            maxLength=20 size=16 name=key style="color: #00006A; font-family: 宋体; font-size: 9pt; background-color: #EBEBEB" value="请输入关键字" onclick="Javascript:this.value='';"><br>
             <br>
             <input type="button" value="搜 索" style="font-family: 宋体; font-size: 9pt; color: #00006A; position: relative; height: 20" onclick="s_check()">
             <br>
             <br>
			</td>
            </tr>
          </table>
          </form>
          </center>
        </div>
        <FORM name=select action="" method=post>
        <p align="center"></p>
        <p align="center">&nbsp;</p>
        </form>
        <p align="center">&nbsp;</p>
        <p>　
      </center>
      </td>
      <td width="22" height="46" valign="top" rowspan="2"><img border="0" src="images/selfk.GIF"></td>
      <td width="446" height="18" valign="top">
      <table border="1" cellpadding="0" cellspacing="0" width="430" height="51" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
        <tr>
          <td width="176" height="1"></td>
          <td width="248" height="10" colspan="2"></td>
        </tr>
        <tr>
            <td width="424" height="20" colspan="3" valign="bottom" background="images/t-bg2.gif"><font color="#000000">&nbsp;<font color="#FFFFFF"> 
              最新五条招聘信息</font></font></td>
        </tr>
        <tr>
          <td width="174" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;公司名称</td>
          <td width="177" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;招聘职位</td>
          <td height="18" bgcolor="#EBEEF3" valign="bottom" width="72">
            <p align="center">发布日期</p>
          </td>
        </tr>
        <% do while not rs.eof %>
        <tr>
  <td width="174" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<a href="javascript:openwin('company.asp?uid=<%=rs("uname")%>','top=10,left=80,width=460,height=420')"><% if len(rs("cname"))>12 then%><%=Left(rs("cname"),12)%>...<% else%><%=rs("cname")%><%end if%></a></td>    
  <td width="172" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<a href="javascript:openwin('job.asp?uid=<%=rs("uname")%>','top=10,left=80,width=460,height=420')"><%=rs("job")%></a></td>
  <td width="72" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center"><%=rs("idate")%></p>
          </td>
        </tr>
     <% c=c+1                                                                     
     rs.movenext                                                                     
     if c>=5 then exit do                                                                     
     loop                                                                    
     rs.close                                                                    
     set rs=nothing %>  
        <tr background="images/t-bg3.gif">
          <td width="424" height="19" colspan="3" valign="bottom">
            <p align="right"><a href="javascript:openwin('search.asp?stype=company&gzdd=noxz','top=10,left=80,width=480,height=350')"><font color="#000000">更多...</font></a></td>
        </tr>
      </table>
  <center>
      <% set rs=server.createobject("adodb.recordset")
	     sql="select * from person" 
         rs.open sql,conn,1,1
         r2=rs.recordcount
         rs.close                                                                    
         set rs=nothing
         set rs=server.createobject("adodb.recordset")
		 sql="select * from person where job<>'""' order by id desc"           
         rs.open sql,conn,1,1 
         jlnum=rs.recordcount %>
      <div align="left">
      <table border="1" cellpadding="0" cellspacing="0" width="430" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
        <tr>
              <td height="20" colspan="5" valign="bottom" background="images/t-bg2.gif" width="426">&nbsp;<font color="#FFFFFF"> 
                最新五条求职信息</font></td>
        </tr>
  </center>
        <tr>
          <td width="67" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="left">&nbsp;姓名</p>   
          </td>
          <td width="45" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">性别</p>  
          </td>
          <td width="59" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">学 历   
          </td>
          <td width="177" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="left">&nbsp;应聘职位</td>
  <center>
          <td width="71" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">登录日期</p>
  </td>
        </tr>
        <% do while not rs.eof %>
      </center>
        <tr>
          <td width="67" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="left">
		   &nbsp;<a href="javascript:openwin('person.asp?uid=<%=rs("uname")%>','top=10,left=80,width=460,height=420')"><%=rs("iname")%></a></p>
          </td>
          <td width="45" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">[<%=rs("sex")%>]</p>
          </td>
  <center>
          <td width="59" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">[<%=rs("edu")%>]</td>
          <td width="176" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<%=rs("job")%></td>
          <td width="71" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center"><%=rs("idate")%></p>
          </td>
        </tr>
     <% p=p+1                                                                     
     rs.movenext                                                                     
     if p>=5 then exit do                                                                     
     loop                                                                    
     rs.close                                                                    
     set rs=nothing 
	 set rs=server.createobject("adodb.recordset")
     sql="select * from pmailbox" 
     rs.open sql,conn,1,1
     pmailnum=rs.recordcount
     rs.close                                                                    
     set rs=nothing 
	 set rs=server.createobject("adodb.recordset")
     sql="select * from cmailbox" 
     rs.open sql,conn,1,1
     cmailnum=rs.recordcount
     rs.close                                                                    
     set rs=nothing %> 
      </center>
        <tr>
          <td width="426" height="19" colspan="5" background="images/t-bg3.gif" valign="bottom">
            <p align="right"><a href="javascript:openwin('search.asp?stype=person&gzdd=noxz','top=10,left=80,width=480,height=350')" ><font color="#000000">更多...</font></a></td>
        </tr>
      </table>
      </div>
      </td>
      <td width="1" height="46" valign="top" rowspan="2" bgcolor="#00006A"></td>
      <td width="150" valign="top" rowspan="2" bgcolor="#F3F3F3" height="46">
        　
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="118" height="260">
            <tr>
              <td height="14" width="117" valign="top">
                <p align="center">=== 站内统计 ===</td>                         
            </tr>
            <tr>
              <td height="3" width="117" valign="top">
              </td>
            </tr>
            <tr>
      <td height="1" valign="top" bgcolor="#000000">
      </td>
            </tr>
            <tr>
              <td height="11" width="117" valign="top">
              </td>
            </tr>
            <tr>
              <td height="54" width="117" valign="top">
                &nbsp;招聘信息:<font color="#0000AA"><%=jobnum%></font>条<br>
                &nbsp;求职简历:<font color="#0000AA"><%=jlnum%></font>份<br>
                &nbsp;注册用户:<font color="#0000AA"><%=r1+r2%></font>位<br>
                &nbsp;站内信件:<font color="#0000AA"><%=pmailnum+cmailnum%></font>封</td>
            </tr>
            <tr>
      <td height="9" valign="top">
      </td>
            </tr>
            <tr>
      <td height="1" valign="top" bgcolor="#000000">
      </td>
            </tr>
            <tr>
      <td height="27" valign="top">
      </td>
            </tr>
            <tr>
      <td height="5" valign="top">
      <p align="center">=== 栏目公告 ===           
      </td>
            </tr>
            <tr>
      <td height="7" valign="top">
      </td>
            </tr>
            <tr>
      <td height="1" valign="top" bgcolor="#000000">
      </td>
            </tr>
            <tr>
      <td height="5" valign="top">
      </td>
            </tr>
            <tr>
      <td height="118" valign="top">
      to inat:<br> 
      &nbsp;&nbsp;&nbsp; 人才市场包括个人登录和单位登录！<br>
      主要功能包括站内信箱（个人用户和单位用户互发站内信件）和我的收藏夹功能！管理部分包括新闻管理和用户管理，以及群发站内信件等...      
      </td>
            </tr>
            <tr>
      <td height="7" valign="top">
      </td>
            </tr>
            <tr>
      <td height="1" valign="top" bgcolor="#000000">
      </td>
            </tr>
          </table>
        </div>
        <div align="center">
          <center>
		  <table border="0" cellpadding="0" cellspacing="0" width="118" height="230">
            <tr>
              <td width="125" height="27" valign="top"></td>
            </tr>
            <tr>
              <td width="125" height="18">
                <p align="center">=== 栏目调查 ===</td>                      
            </tr>
            <tr>
              <td width="125" height="7" valign="top">
                <p align="center"></td>
            </tr>
            <tr>
      <td height="1" valign="top" bgcolor="#000000">
      </td>
            </tr>
            <tr>
              <td width="125" height="6" valign="top"></td>
            </tr>
            <tr>
              <td width="125" height="31" valign="top">&nbsp;&nbsp;               
                您认为本栏目哪些方面急需改进?</td>
            </tr>
            <form name=research action="research.asp" method=post target="_blank" onsubmit="return research_onsubmit()">
            <tr>
              <td width="125" height="69" valign="top"><input type="radio" value="a" name="Options">A.个性化设计<br>
                <input type="radio" value="b" name="Options">B.增加招聘信息<br>
                <input type="radio" value="c" name="Options">C.丰富栏目内容<br>
                <input type="radio" value="d" name="Options">D.改进页面设计<br>
              </td>
            </tr>
            <tr>
      <td height="1" valign="top" bgcolor="#000000">
      </td>
            </tr>
            <tr>
              <td width="125" height="14" valign="top"></td>
            </tr>
            <tr>
              <td width="125" height="25" valign="top">
                <p align="center"><input type="submit" value="提交" style="font-family: 宋体; font-size: 9pt; color: #000060; position: relative; height: 20">&nbsp;           
  <input type="button" value="结果" onclick="javascript:openwin('research.asp?stype=view','top=10,left=80,width=420,height=250')" style="font-family: 宋体; font-size: 9pt; color: #000060; position: relative; height: 20"></td>
            </tr> 
           </form>
          </table>
          </center>
        </div>
      </td>
    </tr>
    <tr>
      
	 <% set rs=server.createobject("adodb.recordset")
     sql="select * from jobnews order by id desc" 
     rs.open sql,conn,1,1 %>
	  <td width="446" height="28" valign="top">
      <div align="left">
        <table border="1" cellpadding="0" cellspacing="0" width="430" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" height="24">
          <tr>
            <td height="18" background="images/t-bg2.gif" valign="bottom" width="426">&nbsp;新闻资讯|NEWS</td>
          </tr>
          <tr>
            
			<td bgcolor="#F0F0F0" valign="bottom">
            <% do while not rs.eof %>
            &nbsp;==&gt;&nbsp;<a href="javascript:openwin('viewnews.asp?id=<%=rs("id")%>','top=20,left=360,width=410,height=350')"><% if len(rs("title"))>18 then%><%=left(rs("title"),15)%>...<% else%><%=rs("title")%><%end if%></a><I> [<%=rs("idate")%>]</I>&nbsp;点击<font color="#000091"><%=rs("click")%></font>次<br>    
            <% n=n+1                                                                     
               rs.movenext                                                                     
               if n>=13 then exit do                                                                     
               loop                                                                    
               rs.close                                                                    
               set rs=nothing %> 
		</td>
		</tr>
          <tr>
          <td height="17" width="424" valign="bottom" background="images/t-bg3.gif">
            <p align="right"><a href="javascript:openwin('jobnews.asp','top=10,left=80,width=500,height=320')"><font color="#000000">更多...</font></a> 
		</td>
          </tr>
        </table>
        </div>
      </td>
    </tr>
    <tr> 
      <td width="778" height="1" colspan="5" valign="top" bgcolor="#53BEB0"> 
        <p align="center"> </td>
    </tr>
    <tr>
      <td height="4" valign="top" colspan="5" width="778">
      <p align="center">
      </td>
    </tr>
    <tr>
      <td height="14" valign="top" colspan="5" width="778">
      <p align="center"><script language="javascript" src="inc/copyright.js"></script>
      </td>
    </tr>
    <tr>
      <td height="3" valign="top" colspan="5" width="778">
      </td>
    </tr>
  </table>
</div>

</body>

</html>