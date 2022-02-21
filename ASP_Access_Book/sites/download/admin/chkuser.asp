<%
	dim Founderr,Errmsg
	dim Founduser
	Founduser=false
	Founderr=false
	dim membername,memberword,memberclass
	dim skin
	membername=request.cookies("aspsky")("username")
	memberword=request.cookies("aspsky")("password")
	memberclass=request.cookies("aspsky")("userclass")
	skin=request.cookies("aspsky")("skin")
	dim guibin,boardmaster,master
	guibin=false
	boardmaster=false
	master=false
	if membername<>"" then
		sql="select userclass from [user] where username='"&membername&"' and userpassword='"&memberword&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			Errmsg=Errmsg+"<br>"+"<li>一般程序保护错误，您试图进行不合法的操作。<li>您的密码不正确，请<a href=logout.asp><font color="&TableContentColor&">退出论坛</font></a>或者<a href=login.asp><font color="&TableContentColor&">重新登陆</font></a>。"
			Founderr=true
		else
			founduser=true
			if rs(0)=18 then
				guibin=true
			elseif rs(0)=19 then
				boardmaster=true
			elseif rs(0)=20 then
				master=true
			else
				founduser=true
			end if
		end if
		rs.close
		set rs=nothing
	end if
	function chkuserlogin(username,password,usercookies,ctype)
	chkuserlogin=false
   	sql="select username,userpassword,lockuser,userclass,article,userWealth,userEP,userCP,usercookies from [User] where username='"&replace(trim(username),"'","''")&"'"
   	set rsUser=server.createobject("adodb.recordset")
	rsUser.open sql,conn,1,3
	if rsUser.eof and rsUser.bof then
		chkuserlogin=false
	else
		if trim(password)<>trim(rsUser(1)) then
			chkuserlogin=false
		elseif rsUser("lockuser")=1 then
			chkuserlogin=false
		else
			chkuserlogin=true
			article=rsUser("article")
			userclass=rsUser("userclass")
			if isnull(userclass) or userclass="" then userclass=1
			if userclass<>18 and userclass<>19 and userclass<>20 then
				if article<cint(point(2)) then userclass=1
				for i=2 to 16
				if article>=point(i) and article<point(i+1) then userclass=i
				next
				if article>=point(17) then userclass=17
			end if
			select case ctype
			case 1
			sql="update [user] set userWealth=userWealth+"&wealthLogin&",userEP=userEP+"&epLogin&",userCP=userCP+"&cpLogin&",userclass="&userclass&",lastlogin=Now(),logins=logins+1 where username='"&replace(trim(username),"'","''")&"'"
			case 2
			sql="update [user] set article=article+1,userWealth=userWealth+"&wealthAnnounce&",userEP=userEP+"&epAnnounce&",userCP=userCP+"&cpAnnounce&",userclass="&userclass&",lastlogin=Now() where username='"&replace(trim(username),"'","''")&"'"
			case 3
			sql="update [user] set article=article+1,userWealth=userWealth+"&wealthreAnnounce&",userEP=userEP+"&epreAnnounce&",userCP=userCP+"&cpreAnnounce&",userclass="&userclass&",lastlogin=Now() where username='"&replace(trim(username),"'","''")&"'"
			end select
			conn.execute(sql)
			if isnull(usercookies) or usercookies="" then usercookies=3
			select case usercookies
	       	 	case 0
  			Response.Cookies("aspsky")("usercookies") = usercookies
	        	case 1
	        	Response.Cookies("aspsky").Expires=Date+1
  			Response.Cookies("aspsky")("usercookies") = usercookies
	        	case 2
	        	Response.Cookies("aspsky").Expires=Date+31
  			Response.Cookies("aspsky")("usercookies") = usercookies
	        	case 3
	        	Response.Cookies("aspsky").Expires=Date+365
  			Response.Cookies("aspsky")("usercookies") = usercookies
	        	end select
			Response.Cookies("aspsky")("username") = rsUser("username")
			Response.Cookies("aspsky")("password") = PassWord
			Response.Cookies("aspsky")("userclass") = grade(userclass)
			rem 清除图片上传数的限制
			response.cookies("aspsky")("upNum")=1
			Response.Cookies("aspsky").path=cookiepath
			call activeuser()
		end if
	end if
	rsUser.close
	set rsUser=nothing
	end function

	function chkboardlogin(boardid,username)
	chkboardlogin=false
	if master then
		chkboardlogin=true
	else
		sql="select boarduser from board where boardid="&boardid
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			chkboardlogin=false
		else
			if isnull(rs(0)) or rs(0)="" then
				chkboardlogin=false
			else
			boarduser=split(rs(0), ",")
			for i = 0 to ubound(boarduser)
				if trim(boarduser(i))=trim(username) then
					chkboardlogin=true
					exit for
				end if
			next
			end if
		end if
		rs.close
		set rs=nothing		
	end if
	end function
	function chkboardmaster(boardid)
	chkboardmaster=false
	if master=true then
		chkboardmaster=true
	else
	sql="select boardmaster,boardid from board where boardid="&boardid
	set brs=server.createobject("adodb.recordset")
	brs.open sql,conn,1,1
	if brs.eof and brs.bof then
		chkboardmaster=false
	else
		if instr(brs("boardmaster"),membername)>0 then
			chkboardmaster=true
		else
			chkboardmaster=false
		end if
	end if
	brs.close
	set brs=nothing
	end if
	end function
%>