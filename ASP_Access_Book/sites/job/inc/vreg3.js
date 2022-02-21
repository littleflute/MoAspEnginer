function check()
{
	if (register.job.value==""){ 
	alert("请选择您的具体应聘岗位！");
		register.job.focus();		
		    return (false);
	}
	if (register.gzdd.value=="") {
		alert("请选择您的希望工作地点！");
		register.gzdd.focus();		
		    return (false);
	}
	if (isNaN(register.yuex.value)){
		alert("月薪栏只能填数字,不填表示面议！");
		register.yuex.focus();		
		    return (false);
    }
	if (isNaN(register.oicq.value)){
		alert("OICQ栏只能填数字！");
		register.oicq.focus();		
		    return (false);
    }
    
    if (document.register.callnum.value.search(/[<>]/gi) != -1) {
    alert("寻呼机号码中包含非法字符<>");
    return false;   
    }
    if (document.register.http.value.search(/[<>]/gi) != -1) {
    alert("个人主页中包含非法字符<>");
    return false;   
    }
	if (checkstring("联系人", document.register.cname.value, false)) {
    document.register.cname.focus();
    return false;   
    }
    if (checkstring("联系电话", document.register.phone.value, false)) {
    document.register.phone.focus();
    return false;   
    } 
	if (checkstring("联系地址", document.register.address.value, false)) {
    document.register.address.focus();
    return false;   
    }
	if (checkemail("电子邮件", document.register.email.value, false)) {
    document.register.email.focus();
    return false;   
    }
	else
		 document.register.submit();
}