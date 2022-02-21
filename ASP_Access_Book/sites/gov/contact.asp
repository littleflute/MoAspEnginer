<!--#include file="rscoon.asp"-->
<!--#include file="totalcount.asp"-->
<% toptitle="留言咨询" %>
<!--#include file="mztop.asp"-->
<script language="javascript">
function form_onsubmit(obj) 
{
  ValidationPassed = true;  
  if(obj.guestname.value=="") {
    alert("请输入名字,以便于我们联系")
    ValidationPassed = false;
    return ValidationPassed;
  }    
  if(obj.from.value == "") {
    alert("为了我们工作的进行,请您填写籍贯,")
    ValidationPassed = false;
    return ValidationPassed;
  }
  if(obj.guestcontent.value == "") {
    alert("请填写你要向我们咨询的内容")
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
              <td width="656" height=21 colspan="2" align=center>请您详细填写您的联系方式,这样才会方便我们尽快给你解决.</td>
          </tr>
          <tr> 
            <td width="656" height="21" align="center"> 
              <div align="right">你的姓名：&nbsp;</div></td>
            <td width="433" height="21" > &nbsp; <input type="text" name="guestname" maxlength="20" size="20" > 
              <font color="#FF0000">*</font> </td>
          </tr>
          <tr> 
            <td width="656" height="21" align="center"> 
              <div align="right">籍贯：&nbsp;</div></td>
            <td width="433" height="21">&nbsp; <input type="text" name="from" maxlength="30" size="30" >
                <font color="#FF0000">*</font> </td>
          </tr>
          <tr> 
            <td width="656" align="center" height="21"> 
              <div align="right">家庭住址：&nbsp;</div></td>
            <td width="433" height="21"> &nbsp; <input name="homeaddr" type="text" size="30" maxlength="30" >
                <font color="#FF0000">*</font> </td>
          </tr>
          <tr> 
            <td width="656" align="center" height="21"> 
              <div align="right">电话号码：&nbsp;</div></td>
            <td width="433" height="21"> &nbsp; <input name="tel" type="text" size="20" maxlength="20" > 
              <font color="#FF0000">*</font> </td>
          </tr>
          <tr> 
            <td width="656" height="22" align="center"> 
              <div align="right">手机号码：&nbsp;</div></td>
            <td width="433" height="22"> &nbsp; <input type="text" name="handset" maxlength="12" size="20">
                (可以不填) </td>
          </tr>
          <tr> 
            <td width="656" height="95" align="center" valign="middle"> 
              <div align="right"> 
                <p style="line-height: 200%">&nbsp;留言内容：(<font color="#FF0000">*</font>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                  <br>
                  &nbsp;&nbsp;&nbsp; <br>
                  &nbsp;&nbsp; </div></td>
            <td width="433" height="95"> 
              &nbsp; <textarea name="guestcontent" cols="47" rows="8"></textarea> 
              <font color="#FF0000">*</font> </td>
          </tr>
          <tr> 
            <td width="656" height=22 colspan="2" align=center> 
              <br> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="submit" name="Submit3" value="提交" onclick="return form_onsubmit(this.form)"> 
              <input type="reset" name="Submit" value="取消"> <br> <br> </td>
          </tr>
        </table>
                  </td>
                </tr>
              </table>
            </form>
<% call mzfoot() %>

