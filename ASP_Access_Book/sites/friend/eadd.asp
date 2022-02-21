<!--#include file="conn.asp"-->
<%
Set rs_lar = Server.CreateObject("ADODB.Recordset")
sql="select * from larchives where user_id like '" & session("user_id") & "'"
rs_lar.open sql,conn,3,2
if rs_lar.eof and rs_lar.bof then
   response.write "您还没建立了档案!"
   response.end
end if
rs_lar.close
name        =request("name")
sex         =request("sex")
byear       =request("byear")
bmonth      =request("bmonth")
bday        =request("bday")
home        =request("home")
education   =request("education")
job         =request("job")
company     =request("company")
postcalcode =request("postcalcode")
tel         =request("tel")
fresume     =request("fresume")
netname     =request("netname")
homepage    =request("homepage")
email       =request("email")
netcall     =request("netcall")
chatroom    =request("chatroom")
sport       =request("sport")
book        =request("book")
music       =request("music")
people      =request("people")
interest    =request("interest")
adage       =request("adage")
character   =request("character")
word        =request("word")
'容错代码--------------------------------------------------------------------------------------------------------------
dim err(26)
if len(name)>10                then err(1)="姓名不能超过10个字符!"
if len(name)<2                 then err(1)="姓名最少也要2个字符!"
if len(home)>30                then err(2)="地/市不能超过30个字符!"
if len(company)>50             then err(3)="单位不能超过50个字符!"
if company=""                  then company="-没填-"
if len(postcalcode)<>6         then err(4)="邮政编码是6位!"
if not(isnumeric(postcalcode)) then err(5)="邮政编码必须是数字!"
if len(tel)>20                 then err(6)="电话号码不能超过20个字符!"
if len(tel)<6                 then err(6)="电话号码不能少于6个字符!"
if not(isnumeric(tel))         then err(7)="电话号码码必须是数字!"
if len(fresume)>100            then err(8)="简历不能超过100个字符!"
if fresume=""                  then fresume="-没填-"
if len(homepage)>50            then err(10)="主页最多只能有50个字符!"
if homepage=""                 then homepage="-没填-"
if len(email)>50               then err(11)="Email最多只能有50个字符!"
if not isvalidemail(email)     then err(12)="Email地址错误!"
if len(chatroom)>50            then err(15)="聊天室地址不能超过50个字符!"
if chatroom=""                 then chatroom="-没填-"
if len(sport)>30               then err(16)="喜欢的运动不能超过30字符!"
if len(book)>50                then err(17)="喜欢的书籍最多不能超过50个字符!"
if len(music)>50               then err(18)="喜欢的音乐最多不能超过50个字符!"
if len(people)>30              then err(19)="喜欢的名人多不能超过50个字符!"
if len(interest)>50            then err(20)="其它爱好最多不能超过50个字符!"
if interest=""                 then interest="-没填-"
if len(adage)>50               then err(21)="喜欢的格言不能超过50个字符!"
if len(character)>50           then err(22)="性格自述不能超过50个字符!"
if len(word)>100               then err(23)="给网友的留言不能超过100个字符!"
if word=""                     then word="-没填-"
if not(isdate(byear & "-" & bmonth & "-" & bday)) then err(24)="生日您没有选择或生日日期无效!"
for i=1 to 25
    if err(i)<>"" then haveerr="yes"
next

function IsValidEmail(email)
 dim names, name, i, c
 IsValidEmail = true
 names = Split(email, "@")
 if UBound(names) <> 1 then
   IsValidEmail = false
   exit function
 end if
 for each name in names
   if Len(name) <= 0 then
     IsValidEmail = false
     exit function
   end if
   for i = 1 to Len(name)
     c = Lcase(Mid(name, i, 1))
     if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
       IsValidEmail = false
       exit function
     end if
   next
   if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
   end if
 next
 if InStr(names(1), ".") <= 0 then
   IsValidEmail = false
   exit function
 end if
 i = Len(names(1)) - InStrRev(names(1), ".")
 if i <> 2 and i <> 3 then
   IsValidEmail = false
   exit function
 end if
 if InStr(email, "..") > 0 then
   IsValidEmail = false
 end if
end function
'--------------------------------------------------------------------------------------------------------------------------

if     (bmonth=12 and bday>=22) or (bmonth=1 and bday<=19)  then
	star= "山羊座"
elseif (bmonth=1 and bday>=20)  or (bmonth=2 and bday<=18)  then
	star= "水瓶座"
elseif (bmonth=2 and bday>=19)  or (bmonth=3 and bday<=20)  then
	star= "双鱼座"
elseif (bmonth=3 and bday>=21)  or (bmonth=4 and bday<=19)  then
	star= "白羊座"
elseif (bmonth=4 and bday>=20)  or (bmonth=5 and bday<=20)  then
	star= "金牛座"
elseif (bmonth=5 and bday>=21)  or (bmonth=6 and bday<=20)  then
	star= "双子座"
elseif (bmonth=6 and bday>=21)  or (bmonth=7 and bday<=22)  then
	star= "巨蟹座"
elseif (bmonth=7 and bday>=23)  or (bmonth=8 and bday<=22)  then
	star= "狮子座"
elseif (bmonth=8 and bday>=23)  or (bmonth=9 and bday<=22)  then
	star= "处女座"
elseif (bmonth=9 and bday>=23)  or (bmonth=10 and bday<=22) then
	star= "天枰座"
elseif (bmonth=10 and bday>=23) or (bmonth=11 and bday<=21) then
	star= "天蝎座"
elseif (bmonth=11 and bday>=22) or (bmonth=12 and bday<=21) then
	star= "射手座"
end if%>

<%if haveerr="yes" then%>
<html>
<head>
<style>
<!--
a:link       { color: blue; text-decoration: none }
a:visited    { color: blue; text-decoration: none }
a:active     { color: #ff9966; text-decoration: none }
a:hover      { color: red; text-decoration: none }
-->
</style>
</head>
<body>
<table border="1" width="400" bordercolor="#CC0000" cellspacing="0" align="center" bgcolor="#B89D77">
  <tr> 
    <td width="100%" align="center"><font color="#FFFFFF"><b>错误提示！</b></font></td>
  </tr>
  <tr>
    <td width="100%">
      <table border="0" width="100%" cellspacing="0" bgcolor="#FFEBE1">
        <tr>
          <td width="100%" align="center">

            <p align="left"><font color="#990000">在您提交的个人档案必栏目中发现如下问题 </font>
          </td>
          </tr>
          <%for i=1 to 24%>
          <%if err(i)<>"" then %>
          <tr>
          <td width="100%">
          <ul>
            <li>
          <font color="red">
          <%=err(i)%>
          </font>
          </li>
          </ul>
          </td>
          </tr>
          <%end if%>
          <%next%>
          <tr>
          <td width="100%" align="center">
          <a href="" onclick="history.go(-1);return false;">[返回]<a>
          </td>
          </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
<%response.end%>
<%end if%>
<%Set rs_lar = Server.CreateObject("ADODB.Recordset")
sql="select * from larchives where user_id like '" & session("user_id") & "'"
rs_lar.open sql,conn,3,2
rs_lar("user_id")      =session("user_id")
rs_lar("name")         =name
rs_lar("sex")          =sex
rs_lar("britherday")   =byear & "-" & bmonth & "-" & bday
rs_lar("age")          =datediff("yyyy",byear & "-" & bmonth & "-" & bday,date)
rs_lar("home")         =home
rs_lar("education")    =education
rs_lar("job")          =job
rs_lar("company")      =company
rs_lar("postcalcode")  =postcalcode
rs_lar("tel")          =tel
rs_lar("fresume")      =fresume
'rs_lar("netname")      =netname
rs_lar("homepage")     =homepage
rs_lar("email")        =email
rs_lar("netcall")      =netcall
rs_lar("chatroom")     =chatroom
rs_lar("sport")        =sport
rs_lar("book")         =book
rs_lar("music")        =music
rs_lar("people")       =people
rs_lar("interest")     =interest
rs_lar("adage")        =adage
rs_lar("character")    =character
rs_lar("ip")           =request.ServerVariables("REMOTE_ADDR")
'rs_lar("date")         =date
'rs_lar("time")         =time
rs_lar("star")         =star
rs_lar.update
rs_lar.close
response.redirect "read.asp?user_id=" & session("user_id")

%>
