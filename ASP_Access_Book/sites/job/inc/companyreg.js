
function checkmod()
{
   if (register.trade.value=="") {
		alert("��ѡ����ҵ������ҵ��");
		register.trade.focus();		
		    return (false);
	}
    if (register.cxz.value=="") {
		alert("��ѡ����ҵ���ʣ�");
		register.cxz.focus();		
		    return (false);
	}
	if (register.area.value=="") {
		alert("��ѡ����ҵ��������");
		register.area.focus();		
		    return (false);
	}
    if (isInvalidDate(register.fdate.value,"-")==true){
    alert("����ȷ��д��˾�ĳ������ڣ�����:1981-11-12����");
    register.fdate.focus();
    return (false);
    }
    if (isNaN(register.zip.value)){
		alert("����ȷ��д�������룡");
		register.zip.focus();		
		    return (false);
    }
     if (isNaN(register.fund.value)){
		alert("����ȷ��дע���ʽ�����");
		register.fund.focus();		
		    return (false);
    }
    if (checkstring("ͨѶ��ַ", document.register.address.value, false)) {
    document.register.address.focus();
    return false;   
    }
    if (register.zip.value=="") {
		alert("�������������룡");
		register.zip.focus();		
		    return (false);
	}
    if (register.zip.value.length!==6) {
		alert("����ȷ��д�������룡");
		register.zip.focus();		
		    return (false);
	}
    if (checkstring("��ϵ������", document.register.pname.value, false)) {
    document.register.pname.focus();
    return false;   
    }
    if (checkstring("��ϵ�绰", document.register.phone.value, false)) {
    document.register.phone.focus();
    return false;   
    }
    if (document.register.fax.value.search(/[<>]/gi) != -1) {
    alert("�����а����Ƿ��ַ�<>");
    return false;   
    }
    if (checkstring("��˾���", document.register.jianj.value, false)) {
    document.register.jianj.focus();
    register.fax.focus();
	return false;   
    }
    if (checkemail("�����ʼ�", document.register.email.value, false)) {
    document.register.email.focus();
    return false;   
    }
    if (document.register.http.value.search(/[<>]/gi) != -1) {
    alert("��˾��ַ�а����Ƿ��ַ�<>");
    register.http.focus();
	return false;   
    }
	else
		 document.register.submit();
}

function check()
{
    if (checkstring("��˾����", document.register.cname.value, false)) {
    document.register.cname.focus();
    return false;   
    }
    checkmod()
}

