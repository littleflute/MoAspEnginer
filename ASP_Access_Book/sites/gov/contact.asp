<!--#include file="rscoon.asp"-->
<!--#include file="totalcount.asp"-->
<% toptitle="������ѯ" %>
<!--#include file="mztop.asp"-->
<script language="javascript">
function form_onsubmit(obj) 
{
  ValidationPassed = true;  
  if(obj.guestname.value=="") {
    alert("����������,�Ա���������ϵ")
    ValidationPassed = false;
    return ValidationPassed;
  }    
  if(obj.from.value == "") {
    alert("Ϊ�����ǹ����Ľ���,������д����,")
    ValidationPassed = false;
    return ValidationPassed;
  }
  if(obj.guestcontent.value == "") {
    alert("����д��Ҫ��������ѯ������")
    ValidationPassed = false;
    return ValidationPassed;
  }
}
</script>
</head>
 <body>
       <form action=contactsave.asp method=post>
              <table cellpadding=0 cellspacing=0 border=0 width=600 align=center >
                <tr> 
                  <td> 
                    <table width=650 height="246" align="center"  cellpadding="0" cellspacing=0 bgcolor="#f1f1eb" style="border-collapse: collapse">
            <tr> 
              <td width="656" height=21 colspan="2" align=center>������ϸ��д������ϵ��ʽ,�����Ż᷽�����Ǿ��������.</td>
          </tr>
          <tr> 
            <td width="656" height="21" align="center"> 
              <div align="right">���������&nbsp;</div></td>
            <td width="433" height="21" > &nbsp; <input type="text" name="guestname" maxlength="20" size="20" > 
              <font color="#FF0000">*</font> </td>
          </tr>
          <tr> 
            <td width="656" height="21" align="center"> 
              <div align="right">���᣺&nbsp;</div></td>
            <td width="433" height="21">&nbsp; <input type="text" name="from" maxlength="30" size="30" >
                <font color="#FF0000">*</font> </td>
          </tr>
          <tr> 
            <td width="656" align="center" height="21"> 
              <div align="right">��ͥסַ��&nbsp;</div></td>
            <td width="433" height="21"> &nbsp; <input name="homeaddr" type="text" size="30" maxlength="30" >
                <font color="#FF0000">*</font> </td>
          </tr>
          <tr> 
            <td width="656" align="center" height="21"> 
              <div align="right">�绰���룺&nbsp;</div></td>
            <td width="433" height="21"> &nbsp; <input name="tel" type="text" size="20" maxlength="20" > 
              <font color="#FF0000">*</font> </td>
          </tr>
          <tr> 
            <td width="656" height="22" align="center"> 
              <div align="right">�ֻ����룺&nbsp;</div></td>
            <td width="433" height="22"> &nbsp; <input type="text" name="handset" maxlength="12" size="20">
                (���Բ���) </td>
          </tr>
          <tr> 
            <td width="656" height="95" align="center" valign="middle"> 
              <div align="right"> 
                <p style="line-height: 200%">&nbsp;�������ݣ�(<font color="#FF0000">*</font>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                  <br>
                  &nbsp;&nbsp;&nbsp; <br>
                  &nbsp;&nbsp; </div></td>
            <td width="433" height="95"> 
              &nbsp; <textarea name="guestcontent" cols="47" rows="8"></textarea> 
              <font color="#FF0000">*</font> </td>
          </tr>
          <tr> 
            <td width="656" height=22 colspan="2" align=center> 
              <br> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="submit" name="Submit3" value="�ύ" onclick="return form_onsubmit(this.form)"> 
              <input type="reset" name="Submit" value="ȡ��"> <br> <br> </td>
          </tr>
        </table>
                  </td>
                </tr>
              </table>
            </form>
<% call mzfoot() %>

