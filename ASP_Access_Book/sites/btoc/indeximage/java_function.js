	function CheckCookie()
	{
	var BOS=document.cookie.indexOf("H19681019S19681231W19980412B69793682=OK")
	
	if (BOS != -1)
	{
		window.open("money.asp",null,"width=560,height=360,scrollbars=1")
	}
	else
	{
		window.open("member.asp",null,"width=500,height=360")
	}
	}