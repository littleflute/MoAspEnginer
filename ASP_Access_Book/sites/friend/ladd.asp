<%
Option Explicit
dim rs_lar
dim sql,haveerr,star
dim name,sex,byear,bmonth,bday,province,home,education,job,company,postcalcode
dim tel,fresume,netname,homepage,email,netcall,netcallcode,chatroom,tallroom,sport,book,music
dim people,interest,adage,character,word
dim i
dim err(26)
%>
<!--#include file="conn.asp"-->
<%
'�Ѷ�Session�����Ƿ�ʱ
if isempty(session("user_id")) then
   response.redirect "timeout.htm"
end if
Set rs_lar = Server.CreateObject("ADODB.Recordset")
sql="select user_id from larchives where user_id =" & session("user_id")
rs_lar.open sql,conn,3,2

if not(rs_lar.eof and rs_lar.bof) then
   response.write "���Ѿ������˵���!"
   response.end
end if

name        =request("name")
sex         =request("sex")
byear       =request("byear")
bmonth      =request("bmonth")
bday        =request("bday")
province    =request("province")
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
netcallcode =request("netcallcode")
chatroom    =request("chatroom")
sport       =request("sport")
book        =request("book")
music       =request("music")
people      =request("people")
interest    =request("interest")
adage       =request("adage")
character   =request("character")
word        =request("word")
'�ݴ����--------------------------------------------------------------------------------------------------------------
if len(name)>10                then err(1)="�������ܳ���10���ַ�!"
if len(name)<2                 then err(1)="��������ҲҪ2���ַ�!"
if len(home)>30                then err(2)="��/�в��ܳ���30���ַ�!"
if len(company)>50             then err(3)="��λ���ܳ���50���ַ�!"
if company=""                  then company="-û��-"
if len(postcalcode)<>6         then err(4)="����������6λ!"
if not(isnumeric(postcalcode)) then err(5)="�����������������!"
if len(tel)>20                 then err(6)="�绰���벻�ܳ���20���ַ�!"
if len(tel)<11                 then err(6)="�绰���벻������11���ַ�!"
if not(isnumeric(tel))         then err(7)="�绰���������������!"
if len(fresume)>100            then err(8)="�������ܳ���100���ַ�!"
if fresume=""                  then fresume="-û��-"
if len(netname)>10             then err(9)="�������ֻ��10���ַ�!"
if len(netname)<2              then err(9)="��������ҲҪ2���ַ�!"
if len(homepage)>50            then err(10)="��ҳ���ֻ����50���ַ�!"
if homepage=""                 then homepage="-û��-"
if len(email)>50               then err(11)="Email���ֻ����50���ַ�!"
if not isvalidemail(email)     then err(12)="Email��ַ����!"
if netcall<>"-û��-" and len(netcallcode)>12 then err(13)="����Ѱ�������벻�ܳ���12���ַ�!"
if netcall<>"-û��-" and len(netcallcode)<6  then err(13)="����Ѱ�������벻������6���ַ�!"
if netcall<>"-û��-" and not isvalidemail(email) then err(14)="����Ѱ��������ֻ��Ϊ����!"
if netcall="-û��-"            then netcallcode="&nbsp;"
if len(chatroom)>50            then err(15)="�����ҵ�ַ���ܳ���50���ַ�!"
if chatroom=""                 then chatroom="-û��-"
if len(sport)>30               then err(16)="ϲ�����˶����ܳ���30�ַ�!"
if sport=""                    then err(16)="ϲ�����˶�����Ϊ��!"
if len(book)>50                then err(17)="ϲ�����鼮��಻�ܳ���50���ַ�!"
if book=""                     then err(17)="ϲ�����鼮����Ϊ��!"
if len(music)>50               then err(18)="ϲ����������಻�ܳ���50���ַ�!"
if music=""                    then err(18)="ϲ�������ֲ���Ϊ��!"
if len(people)>30              then err(19)="ϲ�������˶಻�ܳ���50���ַ�!"
if people=""                   then err(19)="ϲ�������˲���Ϊ��!"
if len(interest)>50            then err(20)="����������಻�ܳ���50���ַ�!"
if interest=""                 then interest="-û��-"
if len(adage)>50               then err(21)="ϲ���ĸ��Բ��ܳ���50���ַ�!"
if adage=""                    then err(21)="ϲ���ĸ��Բ���Ϊ��!"
if len(character)>50           then err(22)="�Ը��������ܳ���50���ַ�!"
if character=""                then err(22)="�Ը��������ܲ���!"
if len(word)>100               then err(23)="�����ѵ����Բ��ܳ���100���ַ�!"
if word=""                     then word="-û��-"
if not(isdate(byear & "-" & bmonth & "-" & bday)) then err(24)="������û��ѡ�������������Ч!"
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
	star= "ɽ����"
elseif (bmonth=1 and bday>=20)  or (bmonth=2 and bday<=18)  then
	star= "ˮƿ��"
elseif (bmonth=2 and bday>=19)  or (bmonth=3 and bday<=20)  then
	star= "˫����"
elseif (bmonth=3 and bday>=21)  or (bmonth=4 and bday<=19)  then
	star= "������"
elseif (bmonth=4 and bday>=20)  or (bmonth=5 and bday<=20)  then
	star= "��ţ��"
elseif (bmonth=5 and bday>=21)  or (bmonth=6 and bday<=20)  then
	star= "˫����"
elseif (bmonth=6 and bday>=21)  or (bmonth=7 and bday<=22)  then
	star= "��з��"
elseif (bmonth=7 and bday>=23)  or (bmonth=8 and bday<=22)  then
	star= "ʨ����"
elseif (bmonth=8 and bday>=23)  or (bmonth=9 and bday<=22)  then
	star= "��Ů��"
elseif (bmonth=9 and bday>=23)  or (bmonth=10 and bday<=22) then
	star= "������"
elseif (bmonth=10 and bday>=23) or (bmonth=11 and bday<=21) then
	star= "��Ы��"
elseif (bmonth=11 and bday>=22) or (bmonth=12 and bday<=21) then
	star= "������"
end if

%>
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
<table border="1" width="400" bordercolor="#336699" cellspacing="0" align="center">
  <tr>
    <td width="100%" bgcolor="#336699" align="center"><font color="#FFFFFF"><b>������ʾ��</b></font></td>
  </tr>
  <tr>
    <td width="100%">
      <table border="0" width="100%" cellspacing="0">
          <tr>
          <td width="100%" align="center">

          <p align="left"><b>�����ύ�ĸ��˵�������Ŀ�з�����������</b>

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
          <a href="" onclick="history.go(-1);return false;">[����]<a>
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
<%

Set rs_lar = Server.CreateObject("ADODB.Recordset")
'������ID,�Ѷ��Ƿ����뵵

rs_lar.open "larchives",conn,3,2
rs_lar.addnew
rs_lar("user_id")      =session("user_id")
rs_lar("name")         =name
rs_lar("sex")          =sex
rs_lar("britherday")   =byear & "-" & bmonth & "-" & bday
rs_lar("age")          =datediff("yyyy",byear & "-" & bmonth & "-" & bday,date)
rs_lar("home")         =province & home
rs_lar("education")    =education
rs_lar("job")          =job
rs_lar("company")      =company
rs_lar("postcalcode")  =postcalcode
rs_lar("tel")          =tel
rs_lar("fresume")      =fresume
rs_lar("netname")      =netname
rs_lar("homepage")     =homepage
rs_lar("email")        =email
if netcall="-û��-" then rs_lar("netcall")="-û��-" else rs_lar("netcall")=netcall & netcallcode
rs_lar("chatroom")     =chatroom
rs_lar("sport")        =sport
rs_lar("book")         =book
rs_lar("music")        =music
rs_lar("people")       =people
rs_lar("interest")     =interest
rs_lar("adage")        =adage
rs_lar("character")    =character
rs_lar("ip")           =request.ServerVariables("REMOTE_ADDR")
rs_lar("date")         =date
'rs_lar("time")         =time
rs_lar("click")        =0
rs_lar("renqi")        =0
rs_lar("photo")        =0
rs_lar("star")         =star
rs_lar.update
rs_lar.close
response.redirect "read.asp?user_id=" & session("user_id")
%>
