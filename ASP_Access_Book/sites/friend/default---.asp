<%Option Explicit%>
<!--#include file="conn.asp"-->
<%
dim connpic
dim rs_user,rs_today,rs_lar,rs_new,rs_board,rs_boy,rs_girl
dim sql,records,today_records,rs_pic
dim i,j,k
 
if isempty(session("user_id")) then session("user_id")=1

'ȡ���������
Set rs_user = Server.CreateObject("ADODB.Recordset")
rs_user.open "user_reg",conn,1,1
records=rs_user.recordcount
rs_user.close:set rs_user=nothing

'ȡ�õ����������
Set rs_today = Server.CreateObject("ADODB.Recordset")
rs_today.open "select * from user_reg where date=" & cdate(date) & "" ,conn,1,1
today_records=rs_today.recordcount
rs_today.close:set rs_today=nothing

'ȡ��������ߵ����ѣ�10λ��
Set rs_lar = Server.CreateObject("ADODB.Recordset")
rs_lar.open "select * from larchives order by renqi desc",conn,1,1

'ȡ��������ߵ�������
Set rs_boy = Server.CreateObject("ADODB.Recordset")
sql="select * from larchives where photo>0 and sex like '" & "��" & "' order by renqi desc"
rs_boy.open sql,conn,1,1

'ȡ��������ߵ�Ů����
Set rs_girl = Server.CreateObject("ADODB.Recordset")
sql="select * from larchives where photo>0 and sex like '" & "Ů" & "' order by renqi desc"
rs_girl.open sql,conn,1,1

'ȡ�����¼�������ѣ�10λ��
Set rs_new = Server.CreateObject("ADODB.Recordset")
rs_new.open "select * from larchives order by lar_id desc",conn,1,1

'ȡ��Ѱ������
Set rs_board = Server.CreateObject("ADODB.Recordset")
rs_board.open "select * from callboard order by id desc",conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Namo WebEditor v4.0(Trial)">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>������վ</title>
<meta name="generator" content="Namo WebEditor v4.0(Trial)">
<style type="text/css"></style>
<link rel="stylesheet" href="ysb.css">
</head>
<body bgcolor="#ffffff" text="black" link="#003300" vlink="#660000" alink="#CC0033" leftmargin="1"
 topmargin="0">
<table width="772" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td><img src="images/top.gif" width="772" height="112" usemap="#Map2" border="0"></td>
  </tr>
</table>
<table border="0" width="772" cellspacing="0" cellpadding="0" bordercolor="white" bordercolordark="white" bordercolorlight="white" bgcolor="#FFFFFF" align="center">
  <tr> 
    <td width="182" bgcolor="#ffffff" valign="top" align="center" bordercolor="#660000"> 
      <table
                                                 border="0" cellspacing="0" width="160"
                                                 height="1" cellpadding="0">
        <tr> 
          <td height="102" bgcolor="white"> 
            <table cellspacing=0 cellpadding=0 width="100%" 
border=0>
              <tbody> 
              <tr> 
                <td width="100%"></td>
              </tr>
              <tr> 
                <td width="100%"><img src="image/03.gif" width="165" height="24"></td>
              </tr>
              <tr> 
                <td> 
                  <div align=center> 
                    <center>
                      <table border="0" width="165" background="image/bg2.gif">
                        <tr align="center" valign="bottom"> 
                          <td width="100%"> 
                            <form action="login.asp" method="POST">
                              <table border="0" width="160" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <td width="31%" align="center" height="25">�û���:</td>
                                  <td width="69%" height="25"> 
                                    <input name=user_name style="BACKGROUND-COLOR:#FFEBE1;BORDER-BOTTOM:#000000 1px solid;BORDER-LEFT:#000000 1px solid;BORDER-RIGHT:#000000 1px solid;BORDER-TOP:#000000 1px solid;FONT-SIZE:9pt;WIDTH:80px">
                                  </td>
                                </tr>
                                <tr valign="bottom"> 
                                  <td width="31%" align="center" height="25">��&nbsp; 
                                    ��:</td>
                                  <td width="69%" height="25"> 
                                    <input name=password style="BACKGROUND-COLOR:#FFEBE1;BORDER-BOTTOM:#000000 1px solid;BORDER-LEFT:#000000 1px solid;BORDER-RIGHT:#000000 1px solid;BORDER-TOP:#000000 1px solid;FONT-SIZE:9pt;WIDTH:80px" type="password">
                                  </td>
                                </tr>
                                <tr align="center" valign="bottom"> 
                                  <td colspan="2" height="30"> 
                                    <input type="image" border="0" name="login" src="images/denglu.GIF" width="30" height="15" alt="�û���¼">
                                    <a href="reg.asp" target="_blank" title="�ո����ֻ꣬�м�ʮ�ſ�Ƭ��"><img src="images/zhuce.GIF" width="30" height="15" border="0" alt="���û�ע��"></a> 
                                  </td>
                                </tr>
                                <tr align="center" valign="bottom"> 
                                  <td colspan="2" height="17">&nbsp; </td>
                                </tr>
                              </table>
                            </form>
                          </td>
                        </tr>
                      </table>
                      <img src="image/02.gif" width="165" height="26" usemap="#Map" border="0"><map name="Map"><area shape="rect" coords="58,4,146,22" href="reg.asp" target="_blank"></map> 
                    </center>
                  </div>
                </td>
              </tr>
              <tr> 
                <td width="100%"></td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
      </table>
      <table border="1" width="160" cellspacing="0" cellpadding="0" bordercolorlight="#330000" bordercolordark="#FFFFFF" bordercolor="#330000">
        <tr bgcolor="#FFCCFF"> 
          <td width="90" height="17" align="center">��Ա����</td>
          <td width="67" height="17" align="center">���ռ���</td>
        </tr>
        <tr bgcolor="#FFEBE1"> 
          <td width="90" align="center" height="20" bgcolor="#FFEBE1"><%=records%><font color="#FF0000">��</font></td>
          <td width="67" align="center" height="20" bgcolor="#FFEBE1"><%=today_records%><font color="#FF0000">��</font></td>
        </tr>
      </table>
      <br>
      <table border="1" width="160" bordercolor="#CC0000" cellspacing="0" cellpadding="0" bordercolorlight="#990000" bordercolordark="#FFFFFF">
        <tr bgcolor="#FFCCFF"> 
          <td width="155" align="center" height="17"> �� �� Ѱ ��</td>
        </tr>
        <tr bgcolor="#FFEBE1"> 
          <td width="155"> 
            <form action="searchout.asp" method="POST">
              <table border="0" width="100%" cellspacing="0" cellpadding="0" bordercolor="#A5EF00">
                <tr> 
                  <td width="15%" height="17"> 
                    <input type="checkbox" name="c_sex" value="ON" checked>
                  </td>
                  <td width="20%" height="17">�Ա�:</td>
                  <td width="65%" height="17"> 
                    <input type="radio" value="��" name="sex">
                    �� 
                    <input type="radio" value="Ů" name="sex" checked>
                    Ů</td>
                </tr>
                <tr>
                  <td width="15%" height="17">
                    <input type="checkbox" name="c_photo" value="ON" checked>
                  </td>
                  <td width="20%" height="17">��Ƭ</td>
                  <td width="65%" height="17">
                    <input type="radio" value="��" name="photo" checked>
                    ��
                    <input type="radio" value="��" name="photo">
                    �� </td>
                </tr>
                <tr> 
                  <td width="15%"> 
                    <input type="checkbox" name="c_name" value="ON">
                  </td>
                  <td width="20%">����:</td>
                  <td width="65%"> 
                    <input name=name style="BACKGROUND-COLOR:#FFEBE1;BORDER-BOTTOM:#000000 1px solid;BORDER-LEFT:#000000 1px solid;BORDER-RIGHT:#000000 1px solid;BORDER-TOP:#000000 1px solid;FONT-SIZE:9pt;WIDTH:90px">
                  </td>
                </tr>
                <tr> 
                  <td width="15%"> 
                    <input type="checkbox" name="c_netname" value="ON">
                  </td>
                  <td width="20%">����:</td>
                  <td width="65%"> 
                    <input name=netname style="BACKGROUND-COLOR:#FFEBE1;BORDER-BOTTOM:#000000 1px solid;BORDER-LEFT:#000000 1px solid;BORDER-RIGHT:#000000 1px solid;BORDER-TOP:#000000 1px solid;FONT-SIZE:9pt;WIDTH:90px" >
                  </td>
                </tr>
                <tr> 
                  <td width="15%"> 
                    <input type="checkbox" name="c_job" value="ON">
                  </td>
                  <td width="20%">ְҵ:</td>
                  <td width="65%"> 
                    <SELECT name="job" size="1">
                      <OPTION selected value=����>����</OPTION>
                      <OPTION value=����>����</OPTION>
                      <OPTION value=���ڲ���>���ڲ���</OPTION>
                      <OPTION value=����>����</OPTION>
                      <OPTION value=����ҵ��>����ҵ��</OPTION>
                      <OPTION value=����>����</OPTION>
                      <OPTION value=����>����</OPTION>
                      <OPTION value=��ʦ>��ʦ</OPTION>
                      <OPTION value=����>����</OPTION>
                      <OPTION value=���>���</OPTION>
                      <OPTION value=����>����</OPTION>
                      <OPTION value=��������>��������</OPTION>
                      <OPTION value=����>����</OPTION>
                      <OPTION value=���ӹ���>���ӹ���</OPTION>
                      <OPTION value=IT>IT</OPTION>
                      <OPTION value=����Ա>����Ա</OPTION>
                      <OPTION value=����>����</OPTION>
                    </SELECT>
                  </td>
                </tr>
                <tr> 
                  <td width="15%"> 
                    <input type="checkbox" name="c_age" value="ON">
                  </td>
                  <td width="20%">����:</td>
                  <td width="65%"> 
                    <input type="text" name="age1" size="2" value="30">
                    �굽 
                    <input type="text" name="age2" size="2" value="20">
                    ��</td>
                </tr>
                <tr valign="bottom"> 
                  <td colspan="3" height="30"> 
                    <p align="center">
                      <input type="submit" name="Submit" value="����">
                    
                  </td>
                </tr>
              </table>
            </form>
          </td>
        </tr>
      </table>
      <br>
    </td>
    <!--1td--> 
    <td width="439" valign="top"> 
      <table border="0" width="100%">
        <tr align="center"> 
          <td> 
            <table border="0" width="419" cellspacing="0" cellpadding="0" height="163">
              <tr> 
                <td width="228" valign="top"> <%Set connpic = Server.CreateObject("ADODB.Connection")    
          DBPath = Server.MapPath("data/picture.mdb")    
          connpic.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath    
		      
          if rs_girl.eof and rs_girl.bof then%> 
                  <table border="0" width="213" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="100" align="right" rowspan="7"><img border="0" src="image/x.jpg" width="79" height="105"></td>
                      <td width="15" align="center" rowspan="7">&nbsp;</td>
                      <td><font color="#FFFFFF">&nbsp;</font></td>
                    </tr>
                    <tr> 
                      <td></td>
                    </tr>
                    <tr> 
                      <td rowspan="5">&nbsp;</td>
                    </tr>
                    <tr> </tr>
                    <tr> </tr>
                    <tr> </tr>
                    <tr> </tr>
                  </table>
                  <%else  
          Set rs_pic = Server.CreateObject("ADODB.Recordset")  
          sql="select * from pic where user_id=" & rs_girl("user_id")  
          rs_pic.open sql,connpic,1,1  
		
          %> 
                  <table border="0" width="209" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="100" align="center" rowspan="7" height="155"><a href='read.asp?user_id=<%=rs_girl("user_id")%>'><font color="black"><img border="1" src='display.asp?id=<%=rs_pic("id")%>' width="90" height="120"></font></a></td>
                      <td width="15" rowspan="7" valign="bottom">&nbsp;</td>
                      <td><font color="black">&nbsp;</font></td>
                    </tr>
		                    <tr> 
                      <td>�Ա�:<font color="#CC0033">Ů</font></td>
                    </tr>
                    <tr> 
                      <td>����:<font color="#ff0000"><%=rs_girl("netname")%></font></td>
                    </tr>
                    <tr> 
                      <td>����:<font color="#ff0000"><%=rs_girl("age")%></font></td>
                    </tr>
                    <tr> 
                      <td>ѧ��:<font color="#ff0000"><%=rs_girl("education")%></font></td>
                    </tr>
                    <tr> 
                      <td>ְҵ:<font color="#ff0000"><%=rs_girl("job")%></font></td>
                    </tr>
                    <tr> 
                      <td>����:<font color="#ff0000"><%=left(rs_girl("home"),3)%></font></td>
                    </tr>
                  </table>
                  <%rs_girl.close:set rs_girl=nothing%> <%end if%> </td>
                <td width="202"> <%if rs_boy.eof and rs_boy.eof then%> 
                  <table border="0" width="206" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="100" align="center" rowspan="7"><img border="0" src="image/x.jpg" width="79" height="105"></td>
                      <td width="15" align="center" rowspan="7">&nbsp;</td>
                      <td width="60">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td width="60"></td>
                    </tr>
                    <tr> 
                      <td width="60"></td>
                    </tr>
                    <tr> 
                      <td width="60"></td>
                    </tr>
                    <tr> 
                      <td width="60"></td>
                    </tr>
                    <tr> 
                      <td width="60" bgcolor="white"></td>
                    </tr>
                    <tr> 
                      <td width="60"></td>
                    </tr>
                  </table>
                  <%else  
          Set rs_pic = Server.CreateObject("ADODB.Recordset")  
          sql="select * from pic where user_id=" & rs_boy("user_id")  
          rs_pic.open sql,connpic,1,1  
          %> 
                  <table border="0" width="208" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="100" align="center" rowspan="7" height="155"><a href='read.asp?user_id=<%=rs_boy("user_id")%>'> 
                        <img border="1" src='display.asp?id=<%=rs_pic("id")%>' width="90" height="120"></a></td>
                      <td width="15" align="center" rowspan="7">&nbsp;</td>
                      <td><font color="black">&nbsp;</font></td>
                    </tr>
                    <tr> 
                      <td>�Ա�:<font color="#FF0000">��</font></td>
                    </tr>
                    <tr> 
                      <td>����:<font color="#ff0000"><%=rs_boy("netname")%></font></td>
                    </tr>
                    <tr> 
                      <td>����:<font color="#ff0000"><%=rs_boy("age")%></font></td>
                    </tr>
                    <tr> 
                      <td>ѧ��:<font color="#ff0000"><%=rs_boy("education")%></font></td>
                    </tr>
                    <tr> 
                      <td>ְҵ:<font color="#ff0000"><%=rs_boy("job")%></font></td>
                    </tr>
                    <tr> 
                      <td>����:<font color="#ff0000"><%=left(rs_boy("home"),3)%></font></td>
                    </tr>
                  </table>
                  <%rs_pic.close:set rs_pic=nothing:set connpic=nothing  
            rs_boy.close:set rs_boy=nothing  
            end if%> </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr align="center"> 
          <td width="100%" valign="top"> 
            <table width="99%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td>&nbsp; </td>
              </tr>
            </table>
            <table width="98%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td> 
                  <table border="0" cellpadding="5" cellspacing="0" align="center" width="100%">
                    <tr> 
                      <td bordercolor="#FF9900" colspan="2"> 
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td width="25%"><img src="image/1%5B3%5D.gif" width="12" height="12"> 
                              ����ǿ���Ƽ�<img src="image/1%5B3%5D.gif" width="12" height="12"></td>
                            <td width="75%"> 
                              <hr noshade size="1">
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                    <tr> 
                      <td align="center" width="30%" rowspan="8"><a href='read.asp?user_id=143'><font color="black"><img border="1" src='display.asp?id=4' width="90" height="120"></font></a><a href="read.asp?user_id=143" target="_blank"><br>
                        ��ϸ����--&gt; </a></td>
                      <td rowspan="8" align="center" valign="top"> 
                        <table width="98%" border="0" cellspacing="0" cellpadding="2">
                          <tr align="center" bgcolor="#FFFFFF"> 
                            <td width="70%" height="15">��1�ڣ�2��12��--3��19�գ�</td>
                          </tr>
                          <tr> 
                            <td width="70%"> 
                              <p>������<font color="#660000"> ��ˮ����</font></p>
                            </td>
                          </tr>
                          <tr align="center"> 
                            <td width="70%" align="left">���ԣ�<font color="#660000"> 
                              �㶫ʡ������</font></td>
                          </tr>
                          <tr align="center"> 
                            <td width="70%" align="left"> 
                              <p>ְҵ�� ����</p>
                            </td>
                          </tr>
                          <tr align="center"> 
                            <td width="70%" align="left"> 
                              <p>�Ը� �Ƚ���ԥ</p>
                            </td>
                          </tr>
                          <tr align="center"> 
                            <td width="70%" align="left"> 
                              <p>OICQ��<font color="#660000"> -û��-</font></p>
                            </td>
                          </tr>
                          <tr align="center"> 
                            <td width="70%" align="left"> 
                              <p>���գ�<font color="#660000"> ����</font></p>
                            </td>
                          </tr>
                          <tr align="center"> 
                            <td width="70%" align="left"> ��ҳ�� <font color="#660000"></font></td>
                          </tr>
                          <tr align="center"> 
                            <td width="70%" align="left"> 
                              <p>���ҽ��ܣ�ϲ������ƽ���</p>
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                    <tr align="center"> </tr>
                    <tr align="center"> </tr>
                    <tr align="center"> </tr>
                    <tr align="center"> </tr>
                    <tr align="center"> </tr>
                    <tr align="center"> </tr>
                    <tr align="center"> </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr align="center"> 
          <td>&nbsp; </td>
        </tr>
      </table>
    </td>
    <!--2td--> 
    <td width="151" bgcolor="white" valign="top" align="center"> 
      <table border="0" width="98%" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
        <tr valign="top"> 
          <td></td><br>
        </tr>
      </table>
      <table border="0" width="100%" bgcolor="#067B0F" cellpadding="1" cellspacing="1">
        <tr> 
          <td width="151"> 
            <table
                                     border="0" cellspacing="0" width="100%" height="1" align="center" cellpadding="0">
              <tr bgcolor="#067B0F"> 
                <td width="153" height="15" bgcolor="#FFCCFF"> 
                  <p
                                                 align="center"><font face="Wingdings" color="#000000">1</font><font face="Wingdings" color="#FFFF33"> 
                    </font><font color="#000000">�������а� </font> 
                </td>
              </tr>
              <tr> 
                <td width="153" height="33" bgcolor="white"> 
                  <table border="0" width="100%" cellspacing="0" cellpadding="0" bgcolor="#FFEBE1">
                    <tr bgcolor="#FFEBE1"> 
                      <td width="28%" align="center" ><font color="#000000">�Ա�</font></td>
                      <td width="46%" align="center" height="16"> 
                        <p align="left"><font color="#000000">����</font></p>
                      </td>
                      <td width="26%" align="center" height="16"><font color="#FFFFFF">����</font> 
                      </td>
                    </tr>
                    <%i=1%> <%do while not(rs_lar.eof) and i<11%> 
                    <tr> 
                      <td width="28%" align="center" height="18" ><%=rs_lar("sex")%></td>
                      <td width="46%" align="center" height="18"> 
                        <p align="left"><a href="read.asp?user_id=<%=rs_lar("user_id")%>"><%=rs_lar("netname")%></a> 
                      </td>
                      <td width="26%" align="center" height="18"><%=rs_lar("renqi")%> 
                      </td>
                    </tr>
                    <%i=i+1%> <%rs_lar.movenext%> <%loop%> <%rs_lar.close%> 
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <br>
      <table border="0" width="100%" bgcolor="#067B0F" cellspacing="1" cellpadding="1">
        <tr> 
          <td width="151"> 
            <table
                                                 border="0" cellspacing="0" width="100%" bordercolor="white" cellpadding="0">
              <tr bgcolor="#067B0F"> 
                <td height="15" bgcolor="#FFCCFF"> 
                  <p
                                                             align="center"><font face="Wingdings" color="#FFFF66">1 
                    </font><font color="#FFFF66">���¼�������</font> 
                </td>
              </tr>
              <tr> 
                <td height="23" bgcolor="white">
                  <table border="0" width="100%" cellspacing="0" cellpadding="0" bgcolor="#FFEBE1">
                    <tr bgcolor="#FFEBE1"> 
                      <td width="26%" bgcolor="#FFEBE1"> 
                        <p align="center"><font color="#000000">�Ա�</font></p>
                      </td>
                      <td width="53%" height="16"> 
                        <p align="left"><font color="#000000">����</font></p>
                      </td>
                      <td width="21%" height="16"><font color="#000000">����</font></td>
                    </tr>
                    <%i=1%> <%do while not(rs_new.eof) and i<11%>
                    <tr bgcolor="#FFEBE1" class="unnamed1">
                      <td width="26%" align="center" height="18"><%=rs_new("sex")%></td>
                      <td width="53%" align="left" height="18" bgcolor="#FFEBE1"> 
                        <p align="left"><a href="read.asp?user_id=<%=rs_new("user_id")%>"><%=rs_new("netname")%></a> 
                      </td>
                      <td width="21%" align="center" height="18"><%=rs_new("renqi")%></td>
                    </tr>
                    <%i=i+1%> <%rs_new.movenext%> <%loop%> <%rs_new.close%>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <br>
      <p>&nbsp;</p>
      <p> [<a href="adminlogin.asp">�������</a>] </p>
      </table>
<table width=772 border=0 cellspacing=0 cellpadding=0 align="center">
  <tr align=center>
    <td colspan=2 height="17"> 
      <hr width=350 size=1 noshade>
  <tr align=center>
    <td colspan=2 height="17"> Copyright &copy;���м���������������޹�˾ 2004


  </table>
<map name="Map2"> 
  <area shape="rect" coords="235,94,274,113" href="default.asp">
  <area shape="rect" coords="298,96,363,111" href="your.asp">
  <area shape="rect" coords="383,95,445,110" href="list.asp">
  <area shape="rect" coords="472,95,531,111" href="register.asp">
  <area shape="rect" coords="556,95,615,109" href="sendphoto.asp">
  <area shape="rect" coords="638,92,698,110" href="pics.asp">
</map>
</html>