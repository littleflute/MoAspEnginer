function checkmod()
{
	if (isInvalidDate(register.bday.value,"-")==true){
    alert("请正确填写您的出生日期（例如:1981-11-12）！");
    register.bday.focus();
    return (false);
    }
	if (isNaN(register.code.value)){
		alert("身份证号码只能填数字！");
		register.code.focus();		
		    return (false);
    }
	if (checkstring("身份证号码", document.register.code.value, false)) {
    document.register.code.focus();
    return false;   
    }
	if ((register.code.value.length!=15) && (register.code.value.length!=18 )) {
		alert("请正确填写您的身份证号码！");
		register.code.focus();		
		    return (false);
	}
   if (checkstring("民族", document.register.mzhu.value, false)) {
    document.register.mzhu.focus();
    return false;   
    }
    if (register.hka.value=="") {
		alert("请选择您的户籍所在地！");
		register.hka.focus();		
		    return (false);
	}
    if (register.edu.value=="") {
		alert("请选择您的最高教育程度！");
		register.edu.focus();		
		    return (false);
	}
    if (checkstring("专业", document.register.zye.value, false)) {
    document.register.zye.focus();
    return false;   
    }
    if (checkstring("毕业院校", document.register.school.value, false)) {
    document.register.school.focus();
    return false;   
    }
    if (register.zzmm.value=="") {
		alert("请选择您的政治面貌！");
		register.zzmm.focus();		
		    return (false);
	}
    if (register.zchen.value=="") {
		alert("请选择您的现有职称！");
		register.zchen.focus();		
		    return (false);
	}
	else
		 document.register.submit();
}

function check()
{
    if (checkstring("真实姓名", document.register.iname.value, false)) {
    document.register.iname.focus();
    return false;   
    }
    checkmod()
}