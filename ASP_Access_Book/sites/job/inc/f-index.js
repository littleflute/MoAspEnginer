function check()
{
    if (login.uname.value=="") 
		alert("请输入用户名！");
	else if (login.pwd.value=="") 
		alert("请输入登录密码！");
	else
		login.submit();
}
function s_check()
{
    if (search.stype.value=="") 
		alert("请选择搜索类别！");
    else if (search.gzdd.value=="") 
		alert("请选择所在地区！");
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
    alert("错误:没有选择！")
    return false
   }
 }