<%
	dim Nstr(20)
	
	Total=Request.Cookies("times") 
	if cstr(Total) = 0 then 
		response.write "���Ĺ��ﳵ��û���κ���Ʒ!"
		response.end
	end if
	for i = 1 to Clng(Total) 
		Nstr(i)=Request.Cookies("MYShopBag"+cstr(i)) 
	next	 

	if Request.Cookies("H19681019S19681231W19980412B69793682") <> "OK" then
		response.redirect "member.asp"
	end if
%>


<html>
<head>
<title>����̨</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="ditu01.gif"><form name="Transfer" method="POST" action="MailToMe.asp">

<table width="520" border="0" cellspacing="0" cellpadding="0" align="center" height="451">
  <tr valign="bottom"> 
      <td width="21" height="38"><img src="image/money01.gif" width="21" height="59"></td>
      <td width="200" background="image/buyditu01.gif" height="38"><img src="image/money02.gif" width="57" height="59"></td>
      <td align="right" background="image/buyditu01.gif" width="418" height="38"><img src="image/money03.gif" width="185" height="35"></td>
      <td align="right" width="35" height="38"><img src="image/buy04.gif" width="35" height="35"></td>
  </tr>
  <tr> 
      <td width="21" background="image/buyditu04.gif" height="342">&nbsp;</td>
    <td colspan="2" valign="top" height="342"> 
      <table width="98%" border="0">
        <tr> 
          <td><font size="2"><span  class="hongzi">1.���������Ĺ��ﶩ��������ϸ�˶���Ʒ���ơ��������۸�</span></font></td>
        </tr>
      </table>
      <table width=98% border=0 cellspacing=1 cellpadding=0 >
        <tr align=center valign=bottom bgcolor="#666666"> 
          <td height=18 width=120><font color="#FFFFFF" size="2">��Ʒ����</font></td>
          <td height=18 width=70><font color="#FFFFFF" size="2">��Ա��</font></td>
          <td height=18 width=60><font color="#FFFFFF" size="2">����</font></td>
            <td height=18 width=80><font color="#FFFFFF" size="2">��Ա��С��</font></td>
        </tr>
        
  <%
  	dim S
  	dim CC
  	
  	for i = 1 to Total 
  		S=split(Nstr(i),"/")
  		CC=CC+Clng(S(1))*Clng(S(2))
  %> 
  
        <tr align=center valign=bottom> 
          <td height=18 width=120><font color="#0000cc" size="2"><%=S(0)%></font></td>
          <td height=18 width=70><font color="#CC0000" size="2"><%=S(1)%></font></td>
          <td height=18 width=60><font size="2"><%=S(2)%></font></td>
          <td height=18 width=80><font color="#CC0000" size="2"><%=cstr(clng(S(1))*clng(S(2)))%></font></td>
        </tr>
       
  <%
    next 
  %>
       
       
      </table>
      <hr width=100%>
      <table width=98% border=0 cellspacing=1 cellpadding=0>
        <tr align=center valign=bottom> 
          <td height=18 width=120>&nbsp;</td>
          <td height=18 width=70>&nbsp;</td>
          <td height=18 width=70>&nbsp;</td>
          <td height=18 width=60><font color="#0000cc" size="2"> �ܼƣ� </font></td>
          <td height=18 width=80><font color="#CC0000" size="2"> <%=CC%>Ԫ</font></td>
        </tr>
      </table>
<p><font  size="2">&nbsp; &nbsp; �������Ҫ��ӡ�˶������뵥������Ҽ����ӵ����Ĳ˵���ѡ�񡰴�ӡ�����<br>&nbsp; &nbsp; �����Ҫ����˶������뵥������Ҽ���ѡ�񡰲鿴Դ�ļ��������ִ�С��ļ���/�����Ϊ��������ļ���� .htm�ļ��ŵ����Ӳ���м��ɡ�</font></p>
    <table width="98%" border="0">
        <tr>
          <td><font size="2">2.�ͻ��ķ�ʽ</font></td>
        </tr>
      </table>
      <table cellspacing=0 cellpadding=0 width=100% border=1 style='font-size: 9pt' bordercolorlight='#D2D2D2' bordercolordark='#FFFFFF' align=center height="42">
        <tr> 
          <td height="18">
            <table border=0 width="442" >
              <tr> 
                <td height="29" width="73" valign="top"> <font size="2"> 
                  <input type='radio' name='Sendit' value='ƽ���ź�' checked >
                  <b>ƽ��</b></font></td>    
                <td height="29" width="355"><font size="2">���ǽ��ԹҺ�ӡˢƷ������ķ�ʽ���͡������ۼ۵Ļ�����,��������ʼġ����������۵Ļ�����,��ȡ���ӷ�<font color="#0000FF">10Ԫ</font>���Һ�1Ԫ������1Ԫ����װ����4Ԫ��ʵ�����ʣ���</font></td>    
              </tr>    
            </table>    
          </td>    
        </tr>    
      </table>    
      <table cellspacing=0 cellpadding=0 width="98%">    
        <tr>     
          <td align=left width="80%" class="hongzi"><font size="2">3.��ѡ�񸶿ʽ</font></td>    
        </tr>    
      </table>    
      <table cellspacing=0 cellpadding=0 width=100% border=1 style='font-size: 9pt' bordercolorlight='#D2D2D2' bordercolordark='#FFFFFF' align=center height="131">    
        <tr>     
          <td height="130">     
            <table border=0  width="445">    
              <tr>     
                <td valign=top width="87" align="left">  
                  <p align="left"> <font size="2">    
                  <b><input type="checkbox" name="C1" value="ON" checked> �����</b></font></p>
                </td>  
                <td width="344"><font size="2"><null> 1.<font color="#0000FF">ͨ���ʾֻ��</font>�������յ���һ��Ҫ3-5��������(ƫԶ�ĵ����������Щ)��<br>
 2.<font color="#0000FF">ͨ�����е��</font>������Ҫ3������������(ƫԶ�ĵ����������Щ)���յ�����,���ǽ��ἰʱ�����������������ʱ������֤��������ƺͶ����ջ�������һ�£���Ϊ��ʱ��ʹ�������˶�����Ҳ�޷��ڵ�㵥����ʾ���������������޷��鵽���Ķ����ţ��ᵢ�����Ķ���������<br>
&nbsp;&nbsp;<font color="#FF0000">ע�⣺</font><font color="#48A289">���������Ļ����ڵ��û����ͻ����ţ���Ҫ��5Ԫ�����ͷѡ�</font>
</font></td>  
              </tr>  
              <tr>   
                <td valign=top width="87"> <font size="2">   
                  <input type='radio' name='Recieve' value='1' checked >  
                  <b>�ʾֻ��</b></font></td>    
                <td width="344"><font size="2">���ڻ���˼��������ע������<font color="#0000FF">������</font>/<font color="#0000FF">�û���</font>������Ҫ�� 
                  <br>
                    ����ַ��������xxxxx<br>
                    �ʱࣺ100000<br>
                    �绰��010-xxxxxx</font></td>    
              </tr>    
              <tr>     
                <td valign=top width="87"> <font size="2">     
                  <input type='radio' name='Recieve' value='2' >    
                  <b>���е��</b></font></td>                             
                <td width="344"><font size="2">��˾��֣������<br>
                    �����У�xxxx֧��<br>
                    �ʺţ�xxxxxxx</font></td>                            
              </tr>                            
            </table>                            
          </td>                            
        </tr>                            
      </table>                            
          </td>                                   
      <td width="35" background="image/buyditu03.gif" height="342">&nbsp;</td>                                   
  </tr>                                   
  <tr valign="top">                                    
      <td width="21" height="27"><img src="image/buy05.gif" width="21" height="21"></td>                  
      <td colspan="2" background="image/buyditu02.gif" height="27">                   
        <p align="right">                  
<input type="button" value="����������?" name="back" onClick="history.go(-1)">&nbsp;<input type="submit" value="��������" name="Send"></p>                     
    </td>                                 
      <td width="35" height="27"><img src="image/buy06.gif" width="35" height="21"></td>                                 
  </tr>                                 
  <tr>                                  
    <td width="21" height="21">&nbsp;</td>                                 
    <td width="200" height="21">&nbsp;</td>                                 
    <td width="418" height="21">&nbsp;</td>                                 
    <td width="35" height="21">&nbsp;</td>                                 
  </tr>                                 
</table>                                 
<table width="520" border="0" cellspacing="0" cellpadding="0" align="center">                          
  <tr>                                 
    <td width="768" height="4" align="center">                                  
      <hr width="100%" size="1" noshade>                                 
    </td>                                 
  </tr>                                 
  <tr>                                  
    <td width="768" height="4" align="center"><font size="2"><b>Copyright</b>:֣������</font></td>                                 
  </tr>                                 
  <tr>                                  
    <td width="768" height="4" align="center"><font size="2"><br>          
      E-mail:<a href="mailto:xxxx@xxxx.com">xxxx@xxxx.com</a></font></td>                              
  </tr>                              
</table>                   
</form>                             
</body>                              
</html>                              
