<html>
<head>
<title>会员身份确认</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="JavaScript">
<!--
function CheckRegister(a,b)
{
	if (a == "" || b == "")
	{
		alert("你必须输入'账号'和'密码'！")
	}
	else
	{
		Register.action="confirm.asp"
		Register.method="POST"
		Register.submit()
	}
}

function CheckAccount(a,b,c,d,e,f,g)
{
	if (a=="" || b=="" || c=="" || d=="" || e=="" || f=="" || g=="")
	{
		alert("请认真填写相应内容,"+"\n" + "可以省却以后的工作")
	}
	else
	{
		if (b!=c)
		{
			alert("两次输入的密码不一致！")
			return false
		}
		else
		{
			Register.action="confirm.asp"
			Register.method="POST"
			Register.submit()
		}
	}
}

//-->


</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" onLoad="MM_preloadImages('indeximage/gouwu02.gif','indeximage/shouyin02.gif','indeximage/chaxun02.gif','indeximage/kefu02.gif')" background="try-image/ditu01.gif" topmargin="0" leftmargin="0">
<form name="Register"><table width="470" border="0" cellspacing="0" cellpadding="0" align="center" height="228">
  <tr valign="bottom"> 
      <td width="22" height="1" ><img src="image/member01.gif" width="21" height="59"></td>
      <td width="83" background="image/buyditu01.gif" height="1" >&nbsp;</td>
      <td align="right" background="image/buyditu01.gif" width="336" height="1"><img src="image/member03.gif" width="240" height="35"></td>
      <td align="right" width="35" height="1"><img src="image/buy04.gif" width="35" height="35"></td>
  </tr>
  <tr> 
    <td width="22" background="image/buyditu04.gif" height="239">&nbsp;</td>
    <td colspan="2" valign="top" height="239" width="421"> 
        <table width="419" border="0" cellspacing="0" cellpadding="2" height="226">
          <tr> 
          <td colspan="3" width="413" height="15"><font size="2"><b>如果您已是洋洋购物的注册会员，请直接填写以下信息：</b></font></td>
        </tr>
        <tr> 
          <td width="128" height="27"><font size="2">账号:                  
            <input type="text" name="A" size="10">                 
            </font></td>                 
          <td width="126" height="27"><font size="2">密码:                  
            <input type="password" name="P" size="10">                 
            </font></td>                 
          <td width="147" height="27">                 
            <input type="button" name="OK" value=" 确 认 " onClick="CheckRegister(Register.A.value,Register.P.value)">                 
          </td>                 
        </tr>                 
        <tr>                  
          <td width="128" height="10">&nbsp;</td>                 
          <td width="126" height="10">&nbsp;</td>                 
          <td width="147" height="10">&nbsp;</td>                 
        </tr>                 
        <tr>                  
          <td colspan="3" width="413" height="21">                  
            <hr width="100%" size="1">                 
          </td>                 
        </tr>                 
        <tr>                  
          <td colspan="3" width="413" height="15"><font size="2"><b>如果你还不是会员并想成为新会员，请填写以下相关内容：</b></font></td>                 
        </tr>                 
        <tr>                  
            <td width="128" bgcolor="#BCBCDE" height="25"><font size="2">&nbsp;账号:        
              <input type="text" name="account" size="9">                 
            </font></td>                 
          <td width="126" height="25"><font size="2">&nbsp;姓名</font><font size="2">:<input type="text" name="myname" size="10">                
            </font></td>                
          <td width="147" height="25"><font size="2">&nbsp;                
            </font><font size="2">电话:<input type="text" name="telephone" size="10">                
            </font></td>                
        </tr>                
        <tr>                 
            <td width="128" bgcolor="#BCBCDE" height="25"><font size="2">&nbsp;密码:        
              <input type="password" name="key1" size="9">                
            </font></td>                
          <td colspan="2" width="273" height="25"><font size="2">&nbsp;地址/邮编:<input type="text" name="address" size="26"></font></td>                
        </tr>                
        <tr>                 
            <td width="128" bgcolor="#BCBCDE" height="25"><font size="2">&nbsp;重输:        
              <input type="password" name="key2" size="9">                 
            </font></td>                 
          <td width="273" colspan="2" height="25"><font size="2">&nbsp;e-Mail</font><font size="1">:</font><font size="2"><input type="text" name="email" size="29"></font></td>                
        </tr>                
        <tr>                 
          <td width="128" height="27">&nbsp; </td>                
          <td width="126" height="27">                
            <p align="right">                
            <input type="button" name="Yes" value=" 提 交 " onClick="           
            	var A=Register.account.value           
            	var B=Register.key1.value           
            	var C=Register.key2.value           
            	var D=Register.address.value           
            	var E=Register.email.value           
            	var F=Register.telephone.value           
            	var G=Register.myname.value           
            	           
            	CheckAccount(A,B,C,D,E,F,G)       	           
				">          
            </p>              
          </td>                
          <td width="147" height="27">                
            &nbsp;                 
            <input type="reset" name="Reset" value=" 重 写 ">                 
          </td>                 
        </tr>                 
      </table>                 
    </td>                 
    <td width="35" background="image/buyditu03.gif" height="239">&nbsp;</td>                 
  </tr>                 
  <tr valign="top">                  
    <td width="22" height="23"><img src="image/buy05.gif" width="21" height="21"></td>                 
    <td colspan="2" background="image/buyditu02.gif" width="421" height="23">&nbsp;</td>                 
    <td width="35" height="23"><img src="image/buy06.gif" width="35" height="21"></td>                 
  </tr>                 
</table>                 
<table width="481" border="0" cellspacing="0" align="center">                 
  <tr>                 
    <td width="749" height="4" align="center">                  
      <hr width="100%" size="1" noshade>                 
    </td>                 
  </tr>                 
  <tr>                  
    <td width="749" height="4" align="center"><font size="2"><b>Copyright</b>:郑州亮中</font></td>                 
  </tr>                 
  <tr>                  
    <td width="749" height="4" align="center"><font size="2"><br> 
      E-mail:<a href="mailto:xxxx@xxxx.com">xxxx@xxxx.com</a></font></td>               
  </tr>               
</table>              
</form>              
<p>&nbsp;</p>               
</body>               
</html>        