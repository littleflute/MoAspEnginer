<% Response.Buffer=True %>
<!--#include file="inc/dbconn.inc"-->
<% 
set rs=server.createobject("adodb.recordset")
sql="select * from company"
rs.open sql,conn,1,1
r1=rs.recordcount
rs.close                                                                    
set rs=nothing
set rs=server.createobject("adodb.recordset")
sql="select * from company where job<>'""' order by id desc"           
rs.open sql,conn,1,1 
jobnum=rs.recordcount %> 

<html>
<title>�����˲š�&gt;��ҳ</title>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" href="inc/index.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

</head>

<body topmargin="0" leftmargin="0">
<!--#include file="inc/top.htm"--> 
<SCRIPT language=JavaScript src="inc/f-index.js"></SCRIPT>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="120">
    <tr bgcolor="#53BEB0"> 
      <td width="778" height="18" colspan="5" valign="top"> �� </td>
    </tr>
    <tr>
      <td height="46" valign="top" bgcolor="#53BEB0" rowspan="2" width="152"> 
        <div align="center">
          <center>
          <table border="0" cellpadding="0" cellspacing="0" width="140" height="142" background="images/loginbg.GIF">
            <tr>
              <td width="136" height="19" valign="top">��</td>
            </tr>
            <form name=login action=login.asp method=post>
			<tr>
              <td width="136" height="47" valign="top">&nbsp;&nbsp;&nbsp;&nbsp; �û���:<input type="text" name="uname" size="7" maxLength=20 style="font-family: ����; font-size: 9pt; background-color: #EBEBEB; color: #00006A"><br>                                                                                                     
                &nbsp;&nbsp;&nbsp;&nbsp; ��&nbsp; ��:<input type="password" name="pwd" maxLength=20 size="7" style="font-family: ����; font-size: 9pt; background-color: #EBEBEB; color: #00006A"></td>                                                                                                      
            </tr>
            <tr>
              <td width="136" height="14" valign="top">&nbsp;&nbsp;&nbsp; <input type="radio" value="person" checked name="usertype">����                                                                                                      
                <input type="radio" name="usertype" value="company">��λ</td>
            </tr>
          </center>
  </center>
            <tr>
              <td width="136" height="71" valign="top">
                <p align="center"><br>
                <input type="button" value="�� ¼" name="B1" style="position: relative; color: #00006A; font-family: ����; font-size: 9pt; height: 19" onClick="check()"><br>
                <br>
                <a href="addnew.asp"><font color="#00006A">���û�ע��</font></a></td>
            </tr>
          </table>
  <center>
          <center></form>
            </center>
        </div>
        <div align="center">
          <center>
          <form name=search action=search.asp method=post target="_blank">
		  <table border="0" cellpadding="0" cellspacing="0" width="138" height="168" background="images/search-bg.GIF">
            <tr>
                      <td width="136" height="167">
                        <p align="center"><font color="#0000FF"><br>
                        </font><font color="#000055">=== վ������ ===</font><font color="#000000"><br>                              
                        <br>
                        </font><SELECT 
                  size=1 name=stype style="background-color: #EBEBEB; color: #00006A; font-family: ����; font-size: 9pt">
                  <OPTION value="" selected>��ѡ���������</option>
				  <option value="company">ְλ����</option>
                  <option value="person">�˲�����</option>
                </SELECT>
                        <p align="center"><SELECT size=1 name=gzdd style="color: #00006A; font-family: ����; font-size: 9pt; background-color: #EBEBEB"> 
              <OPTION selected>��ѡ�����ڵ���</OPTION>
			  <OPTION value=noxz>��&nbsp; ��
			  <OPTION value=������>������
              <OPTION value=�����>�����
              <OPTION value=�Ϻ���>�Ϻ���
			  <OPTION value=������>������
              <OPTION value=����ʡ>����ʡ
              <OPTION value=����ʡ>����ʡ
              <OPTION value=����ʡ>����ʡ
              <OPTION value=�㶫ʡ>�㶫ʡ
			  <OPTION value=�Ĵ�ʡ>�Ĵ�ʡ
              <OPTION value=����ʡ>����ʡ
              <OPTION value=����ʡ>����ʡ
              <OPTION value=����ʡ>����ʡ
			  <OPTION value=����ʡ>����ʡ
              <OPTION value=����ʡ>����ʡ
              <OPTION value=����ʡ>����ʡ
              <OPTION value=����ʡ>����ʡ
              <OPTION value=�ຣʡ>�ຣʡ
			  <OPTION value=ɽ��ʡ>ɽ��ʡ
              <OPTION value=ɽ��ʡ>ɽ��ʡ
			  <OPTION value=����ʡ>����ʡ
              <OPTION value=����ʡ>����ʡ
			  <OPTION value=�㽭ʡ>�㽭ʡ
			  <OPTION value=������ʡ>������ʡ
			  </SELECT></p>
            <p align="center"><INPUT 
            maxLength=20 size=16 name=key style="color: #00006A; font-family: ����; font-size: 9pt; background-color: #EBEBEB" value="������ؼ���" onclick="Javascript:this.value='';"><br>
             <br>
             <input type="button" value="�� ��" style="font-family: ����; font-size: 9pt; color: #00006A; position: relative; height: 20" onclick="s_check()">
             <br>
             <br>
			</td>
            </tr>
          </table>
          </form>
          </center>
        </div>
        <FORM name=select action="" method=post>
        <p align="center"></p>
        <p align="center">&nbsp;</p>
        </form>
        <p align="center">&nbsp;</p>
        <p>��
      </center>
      </td>
      <td width="22" height="46" valign="top" rowspan="2"><img border="0" src="images/selfk.GIF"></td>
      <td width="446" height="18" valign="top">
      <table border="1" cellpadding="0" cellspacing="0" width="430" height="51" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
        <tr>
          <td width="176" height="1"></td>
          <td width="248" height="10" colspan="2"></td>
        </tr>
        <tr>
            <td width="424" height="20" colspan="3" valign="bottom" background="images/t-bg2.gif"><font color="#000000">&nbsp;<font color="#FFFFFF"> 
              ����������Ƹ��Ϣ</font></font></td>
        </tr>
        <tr>
          <td width="174" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;��˾����</td>
          <td width="177" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;��Ƹְλ</td>
          <td height="18" bgcolor="#EBEEF3" valign="bottom" width="72">
            <p align="center">��������</p>
          </td>
        </tr>
        <% do while not rs.eof %>
        <tr>
  <td width="174" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<a href="javascript:openwin('company.asp?uid=<%=rs("uname")%>','top=10,left=80,width=460,height=420')"><% if len(rs("cname"))>12 then%><%=Left(rs("cname"),12)%>...<% else%><%=rs("cname")%><%end if%></a></td>    
  <td width="172" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<a href="javascript:openwin('job.asp?uid=<%=rs("uname")%>','top=10,left=80,width=460,height=420')"><%=rs("job")%></a></td>
  <td width="72" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center"><%=rs("idate")%></p>
          </td>
        </tr>
     <% c=c+1                                                                     
     rs.movenext                                                                     
     if c>=5 then exit do                                                                     
     loop                                                                    
     rs.close                                                                    
     set rs=nothing %>  
        <tr background="images/t-bg3.gif">
          <td width="424" height="19" colspan="3" valign="bottom">
            <p align="right"><a href="javascript:openwin('search.asp?stype=company&gzdd=noxz','top=10,left=80,width=480,height=350')"><font color="#000000">����...</font></a></td>
        </tr>
      </table>
  <center>
      <% set rs=server.createobject("adodb.recordset")
	     sql="select * from person" 
         rs.open sql,conn,1,1
         r2=rs.recordcount
         rs.close                                                                    
         set rs=nothing
         set rs=server.createobject("adodb.recordset")
		 sql="select * from person where job<>'""' order by id desc"           
         rs.open sql,conn,1,1 
         jlnum=rs.recordcount %>
      <div align="left">
      <table border="1" cellpadding="0" cellspacing="0" width="430" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
        <tr>
              <td height="20" colspan="5" valign="bottom" background="images/t-bg2.gif" width="426">&nbsp;<font color="#FFFFFF"> 
                ����������ְ��Ϣ</font></td>
        </tr>
  </center>
        <tr>
          <td width="67" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="left">&nbsp;����</p>   
          </td>
          <td width="45" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">�Ա�</p>  
          </td>
          <td width="59" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">ѧ ��   
          </td>
          <td width="177" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="left">&nbsp;ӦƸְλ</td>
  <center>
          <td width="71" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">��¼����</p>
  </td>
        </tr>
        <% do while not rs.eof %>
      </center>
        <tr>
          <td width="67" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="left">
		   &nbsp;<a href="javascript:openwin('person.asp?uid=<%=rs("uname")%>','top=10,left=80,width=460,height=420')"><%=rs("iname")%></a></p>
          </td>
          <td width="45" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">[<%=rs("sex")%>]</p>
          </td>
  <center>
          <td width="59" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center">[<%=rs("edu")%>]</td>
          <td width="176" height="18" bgcolor="#EBEEF3" valign="bottom">&nbsp;<%=rs("job")%></td>
          <td width="71" height="18" bgcolor="#EBEEF3" valign="bottom">
            <p align="center"><%=rs("idate")%></p>
          </td>
        </tr>
     <% p=p+1                                                                     
     rs.movenext                                                                     
     if p>=5 then exit do                                                                     
     loop                                                                    
     rs.close                                                                    
     set rs=nothing 
	 set rs=server.createobject("adodb.recordset")
     sql="select * from pmailbox" 
     rs.open sql,conn,1,1
     pmailnum=rs.recordcount
     rs.close                                                                    
     set rs=nothing 
	 set rs=server.createobject("adodb.recordset")
     sql="select * from cmailbox" 
     rs.open sql,conn,1,1
     cmailnum=rs.recordcount
     rs.close                                                                    
     set rs=nothing %> 
      </center>
        <tr>
          <td width="426" height="19" colspan="5" background="images/t-bg3.gif" valign="bottom">
            <p align="right"><a href="javascript:openwin('search.asp?stype=person&gzdd=noxz','top=10,left=80,width=480,height=350')" ><font color="#000000">����...</font></a></td>
        </tr>
      </table>
      </div>
      </td>
      <td width="1" height="46" valign="top" rowspan="2" bgcolor="#00006A"></td>
      <td width="150" valign="top" rowspan="2" bgcolor="#F3F3F3" height="46">
        ��
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="118" height="260">
            <tr>
              <td height="14" width="117" valign="top">
                <p align="center">=== վ��ͳ�� ===</td>                         
            </tr>
            <tr>
              <td height="3" width="117" valign="top">
              </td>
            </tr>
            <tr>
      <td height="1" valign="top" bgcolor="#000000">
      </td>
            </tr>
            <tr>
              <td height="11" width="117" valign="top">
              </td>
            </tr>
            <tr>
              <td height="54" width="117" valign="top">
                &nbsp;��Ƹ��Ϣ:<font color="#0000AA"><%=jobnum%></font>��<br>
                &nbsp;��ְ����:<font color="#0000AA"><%=jlnum%></font>��<br>
                &nbsp;ע���û�:<font color="#0000AA"><%=r1+r2%></font>λ<br>
                &nbsp;վ���ż�:<font color="#0000AA"><%=pmailnum+cmailnum%></font>��</td>
            </tr>
            <tr>
      <td height="9" valign="top">
      </td>
            </tr>
            <tr>
      <td height="1" valign="top" bgcolor="#000000">
      </td>
            </tr>
            <tr>
      <td height="27" valign="top">
      </td>
            </tr>
            <tr>
      <td height="5" valign="top">
      <p align="center">=== ��Ŀ���� ===           
      </td>
            </tr>
            <tr>
      <td height="7" valign="top">
      </td>
            </tr>
            <tr>
      <td height="1" valign="top" bgcolor="#000000">
      </td>
            </tr>
            <tr>
      <td height="5" valign="top">
      </td>
            </tr>
            <tr>
      <td height="118" valign="top">
      to inat:<br> 
      &nbsp;&nbsp;&nbsp; �˲��г��������˵�¼�͵�λ��¼��<br>
      ��Ҫ���ܰ���վ�����䣨�����û��͵�λ�û�����վ���ż������ҵ��ղؼй��ܣ������ְ������Ź�����û������Լ�Ⱥ��վ���ż���...      
      </td>
            </tr>
            <tr>
      <td height="7" valign="top">
      </td>
            </tr>
            <tr>
      <td height="1" valign="top" bgcolor="#000000">
      </td>
            </tr>
          </table>
        </div>
        <div align="center">
          <center>
		  <table border="0" cellpadding="0" cellspacing="0" width="118" height="230">
            <tr>
              <td width="125" height="27" valign="top"></td>
            </tr>
            <tr>
              <td width="125" height="18">
                <p align="center">=== ��Ŀ���� ===</td>                      
            </tr>
            <tr>
              <td width="125" height="7" valign="top">
                <p align="center"></td>
            </tr>
            <tr>
      <td height="1" valign="top" bgcolor="#000000">
      </td>
            </tr>
            <tr>
              <td width="125" height="6" valign="top"></td>
            </tr>
            <tr>
              <td width="125" height="31" valign="top">&nbsp;&nbsp;               
                ����Ϊ����Ŀ��Щ���漱��Ľ�?</td>
            </tr>
            <form name=research action="research.asp" method=post target="_blank" onsubmit="return research_onsubmit()">
            <tr>
              <td width="125" height="69" valign="top"><input type="radio" value="a" name="Options">A.���Ի����<br>
                <input type="radio" value="b" name="Options">B.������Ƹ��Ϣ<br>
                <input type="radio" value="c" name="Options">C.�ḻ��Ŀ����<br>
                <input type="radio" value="d" name="Options">D.�Ľ�ҳ�����<br>
              </td>
            </tr>
            <tr>
      <td height="1" valign="top" bgcolor="#000000">
      </td>
            </tr>
            <tr>
              <td width="125" height="14" valign="top"></td>
            </tr>
            <tr>
              <td width="125" height="25" valign="top">
                <p align="center"><input type="submit" value="�ύ" style="font-family: ����; font-size: 9pt; color: #000060; position: relative; height: 20">&nbsp;           
  <input type="button" value="���" onclick="javascript:openwin('research.asp?stype=view','top=10,left=80,width=420,height=250')" style="font-family: ����; font-size: 9pt; color: #000060; position: relative; height: 20"></td>
            </tr> 
           </form>
          </table>
          </center>
        </div>
      </td>
    </tr>
    <tr>
      
	 <% set rs=server.createobject("adodb.recordset")
     sql="select * from jobnews order by id desc" 
     rs.open sql,conn,1,1 %>
	  <td width="446" height="28" valign="top">
      <div align="left">
        <table border="1" cellpadding="0" cellspacing="0" width="430" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" height="24">
          <tr>
            <td height="18" background="images/t-bg2.gif" valign="bottom" width="426">&nbsp;������Ѷ|NEWS</td>
          </tr>
          <tr>
            
			<td bgcolor="#F0F0F0" valign="bottom">
            <% do while not rs.eof %>
            &nbsp;==&gt;&nbsp;<a href="javascript:openwin('viewnews.asp?id=<%=rs("id")%>','top=20,left=360,width=410,height=350')"><% if len(rs("title"))>18 then%><%=left(rs("title"),15)%>...<% else%><%=rs("title")%><%end if%></a><I> [<%=rs("idate")%>]</I>&nbsp;���<font color="#000091"><%=rs("click")%></font>��<br>    
            <% n=n+1                                                                     
               rs.movenext                                                                     
               if n>=13 then exit do                                                                     
               loop                                                                    
               rs.close                                                                    
               set rs=nothing %> 
		</td>
		</tr>
          <tr>
          <td height="17" width="424" valign="bottom" background="images/t-bg3.gif">
            <p align="right"><a href="javascript:openwin('jobnews.asp','top=10,left=80,width=500,height=320')"><font color="#000000">����...</font></a> 
		</td>
          </tr>
        </table>
        </div>
      </td>
    </tr>
    <tr> 
      <td width="778" height="1" colspan="5" valign="top" bgcolor="#53BEB0"> 
        <p align="center"> </td>
    </tr>
    <tr>
      <td height="4" valign="top" colspan="5" width="778">
      <p align="center">
      </td>
    </tr>
    <tr>
      <td height="14" valign="top" colspan="5" width="778">
      <p align="center"><script language="javascript" src="inc/copyright.js"></script>
      </td>
    </tr>
    <tr>
      <td height="3" valign="top" colspan="5" width="778">
      </td>
    </tr>
  </table>
</div>

</body>

</html>