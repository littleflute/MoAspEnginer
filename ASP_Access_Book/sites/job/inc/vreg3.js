function check()
{
	if (register.job.value==""){ 
	alert("��ѡ�����ľ���ӦƸ��λ��");
		register.job.focus();		
		    return (false);
	}
	if (register.gzdd.value=="") {
		alert("��ѡ������ϣ�������ص㣡");
		register.gzdd.focus();		
		    return (false);
	}
	if (isNaN(register.yuex.value)){
		alert("��н��ֻ��������,�����ʾ���飡");
		register.yuex.focus();		
		    return (false);
    }
	if (isNaN(register.oicq.value)){
		alert("OICQ��ֻ�������֣�");
		register.oicq.focus();		
		    return (false);
    }
    
    if (document.register.callnum.value.search(/[<>]/gi) != -1) {
    alert("Ѱ���������а����Ƿ��ַ�<>");
    return false;   
    }
    if (document.register.http.value.search(/[<>]/gi) != -1) {
    alert("������ҳ�а����Ƿ��ַ�<>");
    return false;   
    }
	if (checkstring("��ϵ��", document.register.cname.value, false)) {
    document.register.cname.focus();
    return false;   
    }
    if (checkstring("��ϵ�绰", document.register.phone.value, false)) {
    document.register.phone.focus();
    return false;   
    } 
	if (checkstring("��ϵ��ַ", document.register.address.value, false)) {
    document.register.address.focus();
    return false;   
    }
	if (checkemail("�����ʼ�", document.register.email.value, false)) {
    document.register.email.focus();
    return false;   
    }
	else
		 document.register.submit();
}