function check()
{
    if (login.uname.value=="") 
		alert("�������û�����");
	else if (login.pwd.value=="") 
		alert("�������¼���룡");
	else
		login.submit();
}
function s_check()
{
    if (search.stype.value=="") 
		alert("��ѡ���������");
    else if (search.gzdd.value=="") 
		alert("��ѡ�����ڵ�����");
	else
		search.submit();
}
function go(formobject)
{
if (formobject.options[formobject.selectedIndex].value!="")
{
window.open(formobject.options[formobject.selectedIndex].value)
}
}
function openwin(page,size)
{
window.open(page,"newuser","toolbar=no,location=no,directories=no,status=no,menubar=yes,scrollbars=yes,resizable=no,"+ size);
}                                                                                          
function research_onsubmit() {
 if(document.research.Options[0].checked==false &&document.research.Options[1].checked==false&&document.research.Options[2].checked==false&&document.research.Options[3].checked==false)
  {
    alert("����:û��ѡ��")
    return false
   }
 }