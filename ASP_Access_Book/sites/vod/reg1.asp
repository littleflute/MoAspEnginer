<%PageName="UserReg2"%>
<!--#include file="conn.asp"-->
<!--#include file="topMain.asp"-->
<TITLE>-�û���¼</TITLE>
<table cellpadding=0 cellspacing=0 border=1  width=760  align=center>
  <FORM name=reg2 action=reg2.asp method=post>
    <TR> 
      <td align=center width=100% bgcolor="#3399FF"> 
        <table border="0" cellpadding="0" cellspacing="0" width="317" bordercolor="#3399FF"  height="262">
          <tr> 
            <td width="315" align="right" colspan="2" bgcolor="#3399FF" height="34"> 
              <p align="center">�� 
            </td>
          </tr>
          <tr> 
            <td width="315" align="right" colspan="2" bgcolor="#FFCC99" height="25"> 
              <p align="center">�»�Աע�� 
            </td>
          </tr>
          <tr> 
            <td width="120" align="right" bgcolor="#FFFFFF" height="21"> 
              <p><b><span lang="zh-cn">&nbsp;</span>ע��<span lang="zh-cn">��Ա��</span>��</b> 
            </td>
            <td width="250" bgcolor="#FFFFFF" height="21">&nbsp; 
              <input maxLength="20" class="input1" size="25" name="User">
            </td>
          </tr>
          <tr> 
            <td align="right" width="103" bgcolor="#FFFFFF" height="21"> 
              <p><b><span lang="zh-cn">&nbsp;</span>���룺</b> 
            </td>
            <td width="211" bgcolor="#FFFFFF" height="21">&nbsp; 
              <input type="password" maxLength="20" class="input1" size="25" name="password">
            </td>
          </tr>
          <tr> 
            <td align="right" width="103" bgcolor="#FFFFFF" height="23"> 
              <p><b><span lang="zh-cn">&nbsp;</span>����ȷ�ϣ�</b> 
            </td>
            <td width="211" bgcolor="#FFFFFF" height="23">&nbsp; 
              <input type="password" maxLength="20" class="input1" size="25" name="password2">
            </td>
          </tr>
          <tr> 
            <td align="right" width="103" bgcolor="#FFFFFF" height="22"> 
              <p><b><span lang="zh-cn">&nbsp;��ϵ</span>Email��</b> 
            </td>
            <td width="211" bgcolor="#FFFFFF" height="22">&nbsp; 
              <input maxlength="50" class="input1" size="25" name="Email">
            </td>
          </tr>
          <tr> 
            <td align="right" width="103" bgcolor="#FFFFFF" height="17"> 
              <p><b><span lang="zh-cn">&nbsp;</span>�Ա�</b> 
            </td>
            <td width="211" bgcolor="#FFFFFF" height="17"> 
              <input type="radio" CHECKED value="1" name="sex">
              �к�&nbsp; 
              <input type="radio" value="0" name="sex">
              Ů��</td>
          </tr>
          <tr> 
            <td align="right" width="103" bgcolor="#FFFFFF" height="22"> 
              <p><b><span lang="zh-cn">&nbsp;��ϵ</span>OICQ��</b> 
            </td>
            <td width="211" bgcolor="#FFFFFF" height="22">&nbsp; 
              <input maxLength="20" class="input1" size="25" name="OICQ">
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2" height="19" width="315" bgcolor="#FFFFFF"> 
              <input value="ע ��" name="Submit" type="submit" class="input2" style="height: 26; width:44">
              <input type="reset" value="�� д" name="Submit2" class="input2" style="height: 26 width:44">
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2" height="44" width="315" bgcolor="#3399FF"> ��</td>
          </tr>
        </table>
      </td>
    </tr>
  </FORM>
</table>
<br>
<!--#include file="CopyRight.asp"-->
</HTML>