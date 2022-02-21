<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="author" content="webmaster@fir8.net,Alpar" />
<meta name="keywords" content="Alpar,Alpar's Blog,E'Studio,ASP,Web" />
<meta name="description" content="Alpar's blog - Powered by Alpar" />
<STYLE>TD {
 FONT-SIZE: 9pt; LINE-HEIGHT: 140%
}
body,td,th { color: #0033ff;
	font-size: 11px;
	font-family: Verdana, Arial, Helvetica, sans-serif;
}
a:link,a:visited {
	text-decoration: none;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
}
a:hover {
	text-decoration: none;
	color:#ed1c24;
	font-family: Verdana, Arial, Helvetica, sans-serif;
}
A:active {
 COLOR: #000000; TEXT-DECORATION: none
}
</STYLE></head>
<%
dim nowtime,fyear,fmonth,fday,fweek
nowtime=now()
fyear=year(nowtime)
fmonth=month(nowtime)
fday=day(nowtime)
fweek=weekday(nowtime)
select case fweek
case "1"
  fweek="星期日"
 case "2"
  fweek="星期一"
 case "3"
  fweek="星期二"
 case "4"
  fweek="星期三"
 case "5"
  fweek="星期四"
 case "6"
  fweek="星期五"
 case "7"
  fweek="星期六"
   end select
%>
<% dim mem
Set mem = Conn.exeCute("select * from blog_member order by mem_RegTime desc")
%>
<script>
var marqueeContent = new Array();
marqueeContent[0]='今天是: <%=fyear%> 年 <%=fmonth%> 月 <%=fday%> 日 <%=fweek%>';
marqueeContent[1]='欢迎51pcbook.com的新成员　{<a href="member.asp?action=view&memName=<%=mem("mem_Name")%>"  target="_blank"><%= mem("mem_Name")%></a>} <img src=images/icon_new.gif align="absmiddle">';
<%           
dim rs,Content_Hot
sql=" select top 10 * from blog_Content where log_ViewNums<50000 order by log_ViewNums desc"
set rs = conn.exeCute(sql)
IF NOT (rs.eof AND rs.bof) then
Content_Hot = 2
do while not rs.eof
%>
marqueeContent[<%=Content_Hot%>]='<b><font color=#0033ff>最热日志：</font></b><a href="blogview.asp?logID=<%=rs("log_ID")%>" target="_blank"><%=SplitLines(cutStr(Replace(rs("log_Title"),"<br>",""),47),0)%></a> <img src=images/icon_hot.gif align="absmiddle">  <%=rs("log_PostTime")%>&nbsp;&nbsp;[<%=rs("log_Author")%>]'
<%            rs.movenext
            Content_Hot=Content_Hot+1
        loop
    End If
    rs.close
set rs=nothing
mem.close
set mem=nothing
    conn.close
    set conn=nothing
%>

var marqueeInterval=new Array();  //定义一些常用而且要经常用到的变量
var marqueeId=0;
var marqueeDelay=2000;
var marqueeHeight=17;
//接下来的是定义一些要使用到的函数
function initMarquee() {
    var str=marqueeContent[0];
    document.write('<div id=marqueeBox style="overflow:hidden;height:'+marqueeHeight+'px" onmouseover="clearInterval(marqueeInterval[0])" onmouseout="marqueeInterval[0]=setInterval(\'startMarquee()\',marqueeDelay)"><div>'+str+'</div></div>');
    marqueeId++;
    marqueeInterval[0]=setInterval("startMarquee()",marqueeDelay);
    }
function startMarquee() {
    var str=marqueeContent[marqueeId];
        marqueeId++;
    if(marqueeId>=marqueeContent.length) marqueeId=0;
    if(marqueeBox.childNodes.length==1) {
        var nextLine=document.createElement('DIV');
        nextLine.innerHTML=str;
        marqueeBox.appendChild(nextLine);
        }
    else {
        marqueeBox.childNodes[0].innerHTML=str;
        marqueeBox.appendChild(marqueeBox.childNodes[0]);
        marqueeBox.scrollTop=0;
        }
    clearInterval(marqueeInterval[1]);
    marqueeInterval[1]=setInterval("scrollMarquee()",20);
    }
function scrollMarquee() {
    marqueeBox.scrollTop++;
    if(marqueeBox.scrollTop%marqueeHeight==(marqueeHeight-1)){
        clearInterval(marqueeInterval[1]);
        }
    }
initMarquee();
</script>