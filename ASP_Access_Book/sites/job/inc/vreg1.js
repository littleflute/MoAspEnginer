function checkmod()
{
	if (isInvalidDate(register.bday.value,"-")==true){
    alert("����ȷ��д���ĳ������ڣ�����:1981-11-12����");
    register.bday.focus();
    return (false);
    }
	if (isNaN(register.code.value)){
		alert("���֤����ֻ�������֣�");
		register.code.focus();		
		    return (false);
    }
	if (checkstring("���֤����", document.register.code.value, false)) {
    document.register.code.focus();
    return false;   
    }
	if ((register.code.value.length!=15) && (register.code.value.length!=18 )) {
		alert("����ȷ��д�������֤���룡");
		register.code.focus();		
		    return (false);
	}
   if (checkstring("����", document.register.mzhu.value, false)) {
    document.register.mzhu.focus();
    return false;   
    }
    if (register.hka.value=="") {
		alert("��ѡ�����Ļ������ڵأ�");
		register.hka.focus();		
		    return (false);
	}
    if (register.edu.value=="") {
		alert("��ѡ��������߽����̶ȣ�");
		register.edu.focus();		
		    return (false);
	}
    if (checkstring("רҵ", document.register.zye.value, false)) {
    document.register.zye.focus();
    return false;   
    }
    if (checkstring("��ҵԺУ", document.register.school.value, false)) {
    document.register.school.focus();
    return false;   
    }
    if (register.zzmm.value=="") {
		alert("��ѡ������������ò��");
		register.zzmm.focus();		
		    return (false);
	}
    if (register.zchen.value=="") {
		alert("��ѡ����������ְ�ƣ�");
		register.zchen.focus();		
		    return (false);
	}
	else
		 document.register.submit();
}

function check()
{
    if (checkstring("��ʵ����", document.register.iname.value, false)) {
    document.register.iname.focus();
    return false;   
    }
    checkmod()
}