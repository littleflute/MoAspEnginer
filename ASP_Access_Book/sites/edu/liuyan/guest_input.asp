<%@ language=vbscript %>
<!--#include file="guest_sub.asp"-->
<script language=javascript>
   function Check(theForm){
     if(theForm.UserName.value == "" ||theForm.ly.value=="" ) {
		alert("���ֺ����Բ���Ϊ��,����������д!");
		return false;
	}
	 if(theForm.Email.value.indexOf('@',0)==-1){
	    alert("���E-mail��������,��˶�!")
		return false;
		}
   }
   function popup(name) {                
	window.open("images/" + name , "popup", "width=390,height=300,resizable=0,top=50,left=50");                
   }                
  function init(fields){                
	if( fields.value == fields.defaultValue ) {                
		fields.value = "";               
		location.href = "javascript:void(popup('face/face.htm'))"
	}                
  }               
   function confirm_reset(){
	if (confirm("�������Ҫ���ȫ�������ݣ���ȷ��Ҫ�����?")){
		return true;
	}
	return false;
	}
</script>
<%Header"--��д����"%>
<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;       
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��д����<font color="#FF0000">&nbsp;&nbsp;</font>     
* ע���� Ϊ������ �� ����Բ���&nbsp; *      
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     
<IMG SRC="images/msglist.gif" BORDER="0" WIDTH="16" HEIGHT="21" >&nbsp;<a href="guest_list.asp?PageNo=1">�鿴��������</a>     
</p>     
<form method="POST" action="guest_save.asp" >     
  <input type="hidden" name="form" value="SaveData">     
  <input type="hidden" name="D_Date" value="<%=cstr(now())%>">     
  <p align="center">�� ������ƣ�<font color="#00FF00"><input type="text" name="UserName" size="40" style="background-color: #FFFFFF; border: 1 outset #000000"></font></p>     
  <p align="center">�� �����ʼ���<font color="#00FF00"><input type="text" name="Email" size="40" style="background-color: #FFFFFF; border: 1 outset #000000"></font></p>      
  <p align="center">��       
  ������ҳ��<font color="#00FF00"><input type="text" name="HomePage" size="40" Value="http://" style="background-color: #FFFFFF; border: 1 outset #000000"></font></p>      
  <p align="center">��       
  ���Ժη���<font color="#00FF00"><input type="text" name="LzHf" size="25" style="background-color: #FFFFFF; border: 1 outset #000000">&nbsp;     
  </font>��<A HREF="javascript:void(popup('face/face.htm'))" title=�������鿴Ф����>Ф��</A>      
  <select size="1" name="Face" style="background-color: #FFFFFF; border: 1 outset #000000">      
    <option selected>001</option>
    <option>002</option>
    <option>003</option>
    <option>004</option>
    <option>005</option>
    <option>006</option>
    <option>007</option>
    <option>008</option>
    <option>009</option>
    <option>010</option>
  </select>
  </p>
  <p align="center">��      
  ���Ա��⣺<font color="#00FF00"><input type="text" name="Subject" size="40" style="background-color: #FFFFFF; border: 1 outset #000000"></font></p>      
  <p align="center">��       
  �������ݣ�<font color="#00FF00"><textarea rows="7" name="Ly" cols="40" style="background-color: #FFFFFF; border: 1 outset #000000"></textarea></font></p>      
  <p align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      
  <input type="submit" value="��    ��" onClick="return Check(this.form)" name="B1" style="background-color: #C0C0C0; font-size: 10pt; color: #000000; border: 1 solid #000000; padding-top: 0">&nbsp;       
  <input type="reset" value="��    ��" onClick="return confirm_reset()" name="B2" style="background-color: #C0C0C0; font-size: 10pt; color: #000000; border: 1 solid #000000; padding-top: 0"></p>      
</form>      
<p align="center">��</p>      
<%Bottom%>      
