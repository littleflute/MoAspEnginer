<%
'****************************************************************
'超级管理员身份验证函数

function chkMaster(ChkUser) 
	chkMaster=false
        set chkrs=server.createobject("adodb.recordset")
        sql="select isAdmin from Admin_UserInfo where CategoryName='"&CategoryName&"' and UserName='"&ChkUser&"'"
        chkrs.open sql,conn,1,1 
        if not (chkrs.bof and chkrs.eof) then 
	if chkrs("isAdmin")=True then chkMaster=true		
        end if
	chkrs.close
	set chkrs=nothing
end function

'****************************************************************
'检验用户的发布和管理权限函数，在用户查看权限及管理员设置用户权限时用。

function chk_isPub(ChkUser,SubCateID,CatePub)
	chk_isPub=false
	temps_pub=split(CatePub, ",")
	for i = 0 to ubound(temps_pub)
	if trim(temps_pub(i))=trim(SubCateID) then
		chk_isPub=true
		exit for
        end if
	next
end function

function chk_isAdmin(ChkUser,SubCateID,adminCateAdm)
	chk_isAdmin=false
	temps_Adm=split(adminCateAdm, ",")
	for i = 0 to ubound(temps_Adm)
	if trim(temps_Adm(i))=trim(SubCateID) then
		chk_isAdmin=true
		exit for
        end if
	next
end function

'******************************************************************
'检验用户的发布和管理权限函数，在用户发布和管理文章时用。

function chkPubUser(ChkUser,SubCateID) '检验发布权限
	chkPubUser=false

        set chkrs=server.createobject("adodb.recordset")
        sql="select CatePub,isActive from Admin_UserInfo where CategoryName='"&CategoryName&"' and UserName='"&ChkUser&"'"
        chkrs.open sql,conn,1,1 
        if not (chkrs.bof and chkrs.eof) then 
	temps_pub=split(chkrs("CatePub"), ",")
	for i = 0 to ubound(temps_pub)
	if trim(temps_pub(i))=trim(SubCateID) and chkrs("isActive")=true then
		chkPubUser=true
		exit for
        end if
	next
        end if
	chkrs.close
	set chkrs=nothing
end function


function chkAdmUser(ChkUser,SubCateID) '检验管理权限
	chkAdmUser=false

        set chkrs=server.createobject("adodb.recordset")
        sql="select CateAdm,isActive from Admin_UserInfo where CategoryName='"&CategoryName&"' and UserName='"&ChkUser&"'"
        chkrs.open sql,conn,1,1 
        if not (chkrs.bof and chkrs.eof) then 
	temps_Adm=split(chkrs("CateAdm"), ",")
	for i = 0 to ubound(temps_Adm)
	if trim(temps_Adm(i))=trim(SubCateID) and chkrs("isActive")=true then
		chkAdmUser=true
		exit for
        end if
	next
        end if
	chkrs.close
	set chkrs=nothing
end function


'******************************************************************
'检验用户对软件的修改权。

function chkEditUser(ChkUser,SoftID)
        chkEditUser=false
        set chkrs=server.createobject("adodb.recordset")
        sql="select isAdmin from Admin_UserInfo where CategoryName='"&CategoryName&"' and UserName='"&ChkUser&"'"
        chkrs.open sql,conn,1,1 
        if not (chkrs.bof and chkrs.eof) then 
	if chkrs("isAdmin")=true then chkEditUser=true		
        end if
        chkrs.close
 
         sql="select SoftID from "&CategoryName&"_SoftInfo where username='"&trim(ChkUser)&"' and SoftID="&SoftID
	 set chkrs=server.createobject("adodb.recordset")
	 chkrs.open sql,conn,1,1
         if not (chkrs.bof and chkrs.eof) then chkEditUser=True
	 chkrs.close
	 set chkrs=nothing
end function

'******************************************************************
'日期格式转换函数

function DateTimeFormat(DateTime,Format) 
select case Format
case "1"
 DateTimeFormat=""&year(DateTime)&"年"&month(DateTime)&"月"&day(DateTime)&"日"
case "2"
 DateTimeFormat=""&month(DateTime)&"月"&day(DateTime)&"日"
case "3" 
 DateTimeFormat=""&year(DateTime)&"/"&month(DateTime)&"/"&day(DateTime)&""
case "4"
 DateTimeFormat=""&month(DateTime)&"/"&day(DateTime)&""
case "5"
 DateTimeFormat=""&month(DateTime)&"月"&day(DateTime)&"日 "&FormatDateTime(DateTime,4)&""
case "6"
   temp="周日,周一,周二,周三,周四,周五,周六"
   temp=split(temp,",") 
   DateTimeFormat=temp(Weekday(DateTime)-1)
case else
 DateTimeFormat=DateTime
end select
end function

'******************************************************************
'非法字符过滤函数

function checkStr(str)
	if isnull(str) then
		checkStr = ""
		exit function 
	end if
	checkStr=replace(str,"'","''")
end function


function browser(info)
    if Instr(info,"NetCaptor 6.5.0")>0 then
        browser="NetCaptor 6.5.0"
    elseif Instr(info,"MyIe 3.1")>0 then
        browser="MyIe 3.1"
    elseif Instr(info,"NetCaptor 6.5.0RC1")>0 then
        browser="NetCaptor 6.5.0RC1"
    elseif Instr(info,"NetCaptor 6.5.PB1")>0 then
        browser="NetCaptor 6.5.PB1"
    elseif Instr(info,"MSIE 5.5")>0 then
        browser="Internet Explorer 5.5"
    elseif Instr(info,"MSIE 6.0")>0 then
        browser="Internet Explorer 6.0"
    elseif Instr(info,"MSIE 6.0b")>0 then
        browser="Internet Explorer 6.0b"
    elseif Instr(info,"MSIE 5.01")>0 then
        browser="Internet Explorer 5.01"
    elseif Instr(info,"MSIE 5.0")>0 then
        browser="Internet Explorer 5.00"
    elseif Instr(info,"MSIE 4.0")>0 then
        browser="Internet Explorer 4.01"
    else
        browser="其它"
    end if
end function
function system(info)
    if Instr(info,"NT 5.1")>0 then
        system=system+"Windows XP"
    elseif Instr(info,"Tel")>0 then
        system=system+"Telport"
	elseif Instr(info,"webzip")>0 then
        system=system+"webzip"
	elseif Instr(info,"flashget")>0 then
        system=system+"flashget"
	elseif Instr(info,"offline")>0 then
        system=system+"offline"
    elseif Instr(info,"NT 5")>0 then
        system=system+"Windows 2000"
    elseif Instr(info,"NT 4")>0 then
        system=system+"Windows NT4"
    elseif Instr(info,"98")>0 then
        system=system+"Windows 98"
    elseif Instr(info,"95")>0 then
        system=system+"Windows 95"
	elseif instr(info,"unix") or instr(info,"linux") or instr(info,"SunOS") or instr(info,"BSD") then
	    system=system+"类Unix"
    elseif instr(thesoft,"Mac") then
	    system=system+"Mac"
    else
        system=system+"其它"
    end if
end function

Copyright = Copyright & "<!--" & vbcrlf
Copyright = Copyright & "=================================================" & vbcrlf
Copyright = Copyright & "=     (c)Copyright 2003 ScorpioNet              =" & vbcrlf
Copyright = Copyright & "=	Design		: winner		=" & vbcrlf
Copyright = Copyright & "=	HomePage	:    =" & vbcrlf
Copyright = Copyright & "=	OICQ		: 53854651		=" & vbcrlf
Copyright = Copyright & "=	网址	    : www.hooji.com	=" & vbcrlf
Copyright = Copyright & "=================================================" & vbcrlf
Copyright = Copyright & " -->"
%>