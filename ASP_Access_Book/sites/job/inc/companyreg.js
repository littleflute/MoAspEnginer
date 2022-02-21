
function checkmod()
{
   if (register.trade.value=="") {
		alert("请选择企业所属行业！");
		register.trade.focus();		
		    return (false);
	}
    if (register.cxz.value=="") {
		alert("请选择企业性质！");
		register.cxz.focus();		
		    return (false);
	}
	if (register.area.value=="") {
		alert("请选择企业所在区域！");
		register.area.focus();		
		    return (false);
	}
    if (isInvalidDate(register.fdate.value,"-")==true){
    alert("请正确填写公司的成立日期（例如:1981-11-12）！");
    register.fdate.focus();
    return (false);
    }
    if (isNaN(register.zip.value)){
		alert("请正确填写邮政编码！");
		register.zip.focus();		
		    return (false);
    }
     if (isNaN(register.fund.value)){
		alert("请正确填写注册资金栏！");
		register.fund.focus();		
		    return (false);
    }
    if (checkstring("通讯地址", document.register.address.value, false)) {
    document.register.address.focus();
    return false;   
    }
    if (register.zip.value=="") {
		alert("请输入邮政编码！");
		register.zip.focus();		
		    return (false);
	}
    if (register.zip.value.length!==6) {
		alert("请正确填写邮政编码！");
		register.zip.focus();		
		    return (false);
	}
    if (checkstring("联系人姓名", document.register.pname.value, false)) {
    document.register.pname.focus();
    return false;   
    }
    if (checkstring("联系电话", document.register.phone.value, false)) {
    document.register.phone.focus();
    return false;   
    }
    if (document.register.fax.value.search(/[<>]/gi) != -1) {
    alert("传真中包含非法字符<>");
    return false;   
    }
    if (checkstring("公司简介", document.register.jianj.value, false)) {
    document.register.jianj.focus();
    register.fax.focus();
	return false;   
    }
    if (checkemail("电子邮件", document.register.email.value, false)) {
    document.register.email.focus();
    return false;   
    }
    if (document.register.http.value.search(/[<>]/gi) != -1) {
    alert("公司网址中包含非法字符<>");
    register.http.focus();
	return false;   
    }
	else
		 document.register.submit();
}

function check()
{
    if (checkstring("公司名称", document.register.cname.value, false)) {
    document.register.cname.focus();
    return false;   
    }
    checkmod()
}

