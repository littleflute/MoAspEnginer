<!--#include file="rscoon.asp"-->
<% toptitle="��Ƶ�㲥" %>
<!--#include file="mztop1.asp"-->
<%  dim id
     id=request.querystring("id")  
	 if id="" or isnull(id) or isnumeric(id)<>true then
	 response.write"<script>alert('�Բ���,�������ķǷ�����.\n���β���ȡ��,�뷵������');history.back();</script>"
	  end if
sql="select vdtitle,introvd,movies from video where id="&id
       set rs=conn.execute(sql)
	   if not rs.eof then
%>
<table width="766" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td height="30">
<div align="center"><strong><%=rs("vdtitle")%></strong></div></td>
  </tr>
  <tr>
    <td align="center"> 
      <OBJECT 
                  classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 width=350 height=400 align=middle class=OBJECT 
                  id=MediaPlayer>
        <PARAM NAME="ShowStatusBar" VALUE="-1">
        <PARAM NAME="Filename" VALUE="admingly/movie/<%=rs("movies")%>">
        <embed 
                  src="admingly/movie/<%=rs("movies")%>"
                   width=350 height=400 type=application/x-oleobject 
                  codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 
                  flename=mp></embed></OBJECT>
    </td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp;���ݼ��:&nbsp;<%=rs("introvd")%></td>
  </tr>
  <tr>
    <td><div align="center">Ҫ���߹ۿ���Ƶ�ļ�,��������ð�װ��media player 9.0.�����Ϊ���������ٱȽ���Ҳ�����ڴ����غ�ۿ�<br>
        <strong><font color="#FF0000"><a href="admingly/movie/<%=rs("movies")%>">����˴�����</a></font></strong></div></td>
  </tr>
</table>
<%  else response.write"<script>alert('�벻Ҫ�ڵ�ַ������������Ϣ.\n���β���ȡ��,�뷵������');history.back();</script>"
     end if
 call mzfoot()
%>

