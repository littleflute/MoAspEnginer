<% Response.Buffer=True %>
<!--#include file="../inc/company.inc"-->
<!--#include file="../inc/html.inc"-->
<% uname=session("cuid") 
   modify=request("modify")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from company where uname='"&uname&"'and cname<>'""'"
rs.open sql,conn,1,1 %>

<% if rs.eof or rs.bof then
   response.write"<SCRIPT language=JavaScript>alert('����δ��¼��˾���ϣ����ȵ�¼��˾���ϣ�');"
   response.write"javascript:history.go(-1);</SCRIPT>" 
   end if %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/register.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�����˲š�&gt;������Ƹ��Ϣ</title>
</head>
<SCRIPT LANGUAGE="JavaScript">
<!--//
function check()
{
  if (register.job.value=="") {
		alert("��ѡ����Ƹְλ���ƣ�");
		register.job.focus();		
		    return (false);
  }
  if (register.zpnum.value=="") {
		alert("��������Ƹ������");
		register.zpnum.focus();		
		    return (false);
  }
  if (isNaN(register.zpnum.value)){
		alert("����ȷ��д��Ƹ������");
		register.zpnum.focus();		
		    return (false);
  }
  if (register.gzdd.value=="") {
		alert("��ѡ�����ص㣡");
		register.gzdd.focus();		
		    return (false);
  }
   if (register.zptext.value=="") {
		alert("�������λ������");
		register.zptext.focus();		
		    return (false);
  }
   if (register.xgyq.value=="") {
		alert("���������Ҫ��");
		register.xgyq.focus();		
		    return (false);
  }
else
	document.register.submit();
}
//-->
</SCRIPT>
<FORM name=register action=publish.asp method=post>
<body topmargin="0" leftmargin="0">
<!--#include file="../inc/top2.htm"--> 
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="87">
    <tr>
      <td width="780" height="18" valign="top" colspan="5" bgcolor="#53BEB0"> 
        �� </td>
    </tr>
    <tr>
      <td width="139" height="162" valign="top" bgcolor="#53BEB0"> �� 
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="118" height="280">
            <tr>
              <td width="100%" height="163" background="../images/stat-bg.GIF" valign="top">
        <p align="center"><br>
        <a href="main.asp">��¼��ҳ</a><br>
        <br>
		<a href="register.asp">��¼/���¹�˾����</a><br>
        <br>
		<a href="publish.asp">����/������Ƹ��Ϣ</a><br>
        <br>
        <a href="../changepwd.asp?stype=company" target="_blank">�޸ĵ�¼����</a><br>
        <br>
        <a href="search.asp">ȫ���˲��б�</a>
              </td>
            </tr>
            <tr>
              <td width="100%" height="117" background="../images/stat-bg.GIF" valign="top">
        <p align="center"><br>
        <br>
        <a href="favorite.asp">�ҵ��ղؼ�<br>
        <br>
        <a href="mailbox.asp">�ҵ�����</a><br>
        <br>
        <a href="../exit.asp">�˳���¼</a>
              </td>
            </tr>
          </table>
        </div>
        <p align="center"> </td>
      <td width="37" height="351" valign="top"><img border="0" src="../images/selfk.GIF"></td>
  </center>
      <td width="484" height="351">
        ��
        <div align="left">
        <table border="1" cellpadding="0" cellspacing="0" width="92%" height="312" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
          <tr>
            <td width="100%" height="18" valign="bottom" bgcolor="#C6CEDE" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
              <p align="center">=== ������Ƹ��Ϣ ===</td>                               
          </tr>
          <tr>
            <td width="100%" height="305" valign="top" bgcolor="#F3F3F3">
              <p align="center">��<br>
              ��Ƹְλ:<select size="1" name="job"><OPTION>���������б���ѡ��</OPTION> 
			  <OPTION>-----------��Ӫ/������-----------
			  <OPTION value=(��/��)�ܲ�/�ܾ���/CEO <%if rs("job") ="(��/��)�ܲ�/�ܾ���/CEO" then Response.Write "selected"%>>(��/��)�ܲ�/�ܾ���/CEO
			  <OPTION value=��ҵ�߻� <%if rs("job") ="��ҵ�߻�" then Response.Write "selected"%>>��ҵ�߻�
			  <OPTION value=Ͷ�ʹ��� <%if rs("job") ="Ͷ�ʹ���" then Response.Write "selected"%>>Ͷ�ʹ���
			  <OPTION value=��ܲ�����/��ҵ������� <%if rs("job") ="��ܲ�����/��ҵ�������" then Response.Write "selected"%>>��ܲ�����/��ҵ�������
			  <OPTION value=��ҵ������Ա <%if rs("job") ="��ҵ������Ա" then Response.Write "selected"%>>��ҵ������Ա
	          <OPTION value=�ʹܲ�����/����������� <%if rs("job") ="�ʹܲ�����/�����������" then Response.Write "selected"%>>�ʹܲ�����/�����������
			  <OPTION value=�ʹ���Ա/��������ʦ <%if rs("job") ="�ʹ���Ա/��������ʦ" then Response.Write "selected"%>>�ʹ���Ա/��������ʦ
			  <OPTION value=��������/��Ŀ����/CTO <%if rs("job") ="��������/��Ŀ����/CTO" then Response.Write "selected"%>>��������/��Ŀ����/CTO
			  <OPTION value=��Ϣ����/CIO <%if rs("job") ="��Ϣ����/CIO" then Response.Write "selected"%>>��Ϣ����/CIO
			  <OPTION>--------------������--------------
			  <OPTION value=�����ܼ�/����/������/���� <%if rs("job") ="�����ܼ�/����/������/����" then Response.Write "selected"%>>�����ܼ�/����/������/����
			  <OPTION value=���ʦ <%if rs("job") ="���ʦ" then Response.Write "selected"%>>���ʦ<OPTION value=���>���
			  <OPTION value=����/���� <%if rs("job") ="����/����" then Response.Write "selected"%>>����/����<OPTION value=ͳ�� <%if rs("job") ="ͳ��" then Response.Write "selected"%>>ͳ��
			  <OPTION value=��� <%if rs("job") ="���" then Response.Write "selected"%>>���
			  <OPTION>-----------����/ҵ����-----------
			  <OPTION value=���۾���/������/���� <%if rs("job") ="���۾���/������/����" then Response.Write "selected"%>>���۾���/������/����
			  <OPTION value=����/ó��/����ҵ�� <%if rs("job") ="����/ó��/����ҵ��" then Response.Write "selected"%>>����/ó��/����ҵ��
			  <OPTION value=���۹���ʦ <%if rs("job") ="���۹���ʦ" then Response.Write "selected"%>>���۹���ʦ
			  <OPTION value=ҵ��Ա/ҵ����� <%if rs("job") ="ҵ��Ա/ҵ�����" then Response.Write "selected"%>>ҵ��Ա/ҵ�����
			  <OPTION value=���� <%if rs("job") ="����" then Response.Write "selected"%>>����
			  <OPTION>--------------�г���--------------
			  <OPTION value=�г�����/������ <%if rs("job") ="�г�����/������" then Response.Write "selected"%>>�г�����/������
			  <OPTION value=�г�/�����߻� <%if rs("job") ="�г�/�����߻�" then Response.Write "selected"%>>�г�/�����߻�
			  <OPTION value=�г�����/ҵ����� <%if rs("job") ="�г�����/ҵ�����" then Response.Write "selected"%>>�г�����/ҵ�����
			  <OPTION value=���ء����������� <%if rs("job") ="���ء�����������" then Response.Write "selected"%>>���ء�����������
			  <OPTION>--------------�����--------------
			  <OPTION value=����/ͼ����� <%if rs("job") ="����/ͼ�����" then Response.Write "selected"%>>����/ͼ�����
			  <OPTION value=������ <%if rs("job") ="������" then Response.Write "selected"%>>������
			  <OPTION value=�İ��߻� <%if rs("job") ="�İ��߻�" then Response.Write "selected"%>>�İ��߻�
			  <OPTION value=��ҵ���/��Ʒ��� <%if rs("job") ="��ҵ���/��Ʒ���" then Response.Write "selected"%>>��ҵ���/��Ʒ���
			  <OPTION value=��ý����������� <%if rs("job") ="��ý�����������" then Response.Write "selected"%>>��ý�����������
			  <OPTION value=װ����� <%if rs("job") ="װ�����" then Response.Write "selected"%>>װ�����
			  <OPTION value=����Ʒ��� <%if rs("job") ="����Ʒ���" then Response.Write "selected"%>>����Ʒ���
			  <OPTION value=��֯��װ��� <%if rs("job") ="��֯��װ���" then Response.Write "selected"%>>��֯��װ���
			  <OPTION value=�Ҿ�/�鱦��� <%if rs("job") ="�Ҿ�/�鱦���" then Response.Write "selected"%>>�Ҿ�/�鱦���
			  <OPTION value=���Ի�ͼ��Ա <%if rs("job") ="���Ի�ͼ��Ա" then Response.Write "selected"%>>���Ի�ͼ��Ա
			  <OPTION>------------�ͻ�������------------
			  <OPTION value=�ͷ�������/������ <%if rs("job") ="�ͷ�������/������" then Response.Write "selected"%>>�ͷ�������/������
			  <OPTION value=����֧��/�ͻ���ѵ <%if rs("job") ="����֧��/�ͻ���ѵ" then Response.Write "selected"%>>����֧��/�ͻ���ѵ
			  <OPTION value=�ͷ�/������ѯ <%if rs("job") ="�ͷ�/������ѯ" then Response.Write "selected"%>>�ͷ�/������ѯ
			  <OPTION value=ǰ̨/�Ӵ� <%if rs("job") ="ǰ̨/�Ӵ�" then Response.Write "selected"%>>ǰ̨/�Ӵ�
			  <OPTION>-----------����/������-----------
			  <OPTION value=����/������Դ���� <%if rs("job") ="����/������Դ����" then Response.Write "selected"%>>����/������Դ����
			  <OPTION value=����/������Ա <%if rs("job") ="����/������Ա" then Response.Write "selected"%>>����/������Ա
			  <OPTION value=Ա����ѵ��Ա <%if rs("job") ="Ա����ѵ��Ա" then Response.Write "selected"%>>Ա����ѵ��Ա
			  <OPTION value=��ҵ�Ļ�/���� <%if rs("job") ="��ҵ�Ļ�/����" then Response.Write "selected"%>>��ҵ�Ļ�/����
			  <OPTION>--------------��ְ��--------------
			  <OPTION value=Ӣ�﷭�� <%if rs("job") ="Ӣ�﷭��" then Response.Write "selected"%>>Ӣ�﷭��
			  <OPTION value=�������﷭�� <%if rs("job") ="�������﷭��" then Response.Write "selected"%>>�������﷭��
			  <OPTION value=ͼ���鱨/���Ϲ��� <%if rs("job") ="ͼ���鱨/���Ϲ���" then Response.Write "selected"%>>ͼ���鱨/���Ϲ���
			  <OPTION value=�������ϱ�д <%if rs("job") ="�������ϱ�д" then Response.Write "selected"%>>�������ϱ�д
			  <OPTION value=����/�߼���Ա <%if rs("job") ="����/�߼���Ա" then Response.Write "selected"%>>����/�߼���Ա
			  <OPTION value=��Ա/���Դ���Ա/����Ա <%if rs("job") ="��Ա/���Դ���Ա/����Ա" then Response.Write "selected"%>>��Ա/���Դ���Ա/����Ա
			  <OPTION>-----------��ҵ/������-----------
			  <OPTION value=����/������ <%if rs("job") ="����/������" then Response.Write "selected"%>>����/������
			  <OPTION value=�������� <%if rs("job") ="��������" then Response.Write "selected"%>>��������
			  <OPTION value=���̹��� <%if rs("job") ="���̹���" then Response.Write "selected"%>>���̹���
			  <OPTION value=Ʒ�ʹ��� <%if rs("job") ="Ʒ�ʹ���" then Response.Write "selected"%>>Ʒ�ʹ���
			  <OPTION value=���Ϲ��� <%if rs("job") ="���Ϲ���" then Response.Write "selected"%>>���Ϲ���
			  <OPTION value=�豸���� <%if rs("job") ="�豸����" then Response.Write "selected"%>>�豸����
			  <OPTION value=�ɹ����� <%if rs("job") ="�ɹ�����" then Response.Write "selected"%>>�ɹ�����
			  <OPTION value=�ֿ���� <%if rs("job") ="�����" then Response.Write "selected"%>>�ֿ����
			  <OPTION value=�ƻ�Ա <%if rs("job") ="�ƻ�Ա" then Response.Write "selected"%>>�ƻ�Ա
			  <OPTION value=���鹤�� <%if rs("job") ="���鹤��" then Response.Write "selected"%>>���鹤��
			  <OPTION value=���� <%if rs("job") ="����" then Response.Write "selected"%>>����
			  <OPTION value=�չ� <%if rs("job") ="�չ�" then Response.Write "selected"%>>�չ�
			  <OPTION>-----------����/������-----------
			  <OPTION value=˾�� <%if rs("job") ="˾��" then Response.Write "selected"%>>˾��
			  <OPTION value=����/��ʦ/��๤ <%if rs("job") ="����/��ʦ/��๤ " then Response.Write "selected"%>>����/��ʦ/��๤
			  <OPTION value=����Ա <%if rs("job") ="����Ա " then Response.Write "selected"%>>����Ա
			  <OPTION value=ӪҵԱ <%if rs("job") ="ӪҵԱ " then Response.Write "selected"%>>ӪҵԱ
			  <OPTION>----------�����רҵ��Ա----------
			  <OPTION value=ϵͳ����Ա <%if rs("job") ="ϵͳ����Ա" then Response.Write "selected"%>>ϵͳ����Ա
			  <OPTION value=��������ԣ�����ʦ <%if rs("job") ="��������ԣ�����ʦ " then Response.Write "selected"%>>��������ԣ�����ʦ
			  <OPTION value=Ӳ�������ԣ�����ʦ <%if rs("job") ="Ӳ�������ԣ�����ʦ" then Response.Write "selected"%>>Ӳ�������ԣ�����ʦ
			  <OPTION value=ϵͳ����ʦ/���� <%if rs("job") ="ϵͳ����ʦ/����" then Response.Write "selected"%>>ϵͳ����ʦ/����
			  <OPTION value=��վ�߻�/��Ϣ�༭ <%if rs("job") ="��վ�߻�/��Ϣ�༭" then Response.Write "selected"%>>��վ�߻�/��Ϣ�༭
			  <OPTION value=��ҳ���/�������� <%if rs("job") ="��ҳ���/��������" then Response.Write "selected"%>>��ҳ���/��������
			  <OPTION value=Internet/WEB/�������񿪷� <%if rs("job") ="Internet/WEB/�������񿪷�" then Response.Write "selected"%>>Internet/WEB/�������񿪷�
			  <OPTION>-------����/ͨѶ��רҵ��Ա-------
			  <OPTION value=���ӹ���ʦ <%if rs("job") ="���ӹ���ʦ" then Response.Write "selected"%>>���ӹ���ʦ
			  <OPTION value=����Ԫ��������ʦ <%if rs("job") ="����Ԫ��������ʦ" then Response.Write "selected"%>>����Ԫ��������ʦ
			  <OPTION value=�Զ����� <%if rs("job") ="�Զ�����" then Response.Write "selected"%>>�Զ�����
			  <OPTION value=���ܴ���/�ۺϲ���/���� <%if rs("job") ="���ܴ���/�ۺϲ���/����" then Response.Write "selected"%>>���ܴ���/�ۺϲ���/����
			  <OPTION value=�����Ǳ�/���� <%if rs("job") ="�����Ǳ�/����" then Response.Write "selected"%>>�����Ǳ�/����
			  <OPTION value=���� <%if rs("job") ="����" then Response.Write "selected"%>>����
			  <OPTION value=���� <%if rs("job") ="����" then Response.Write "selected"%>>����
			  <OPTION value=ͨѶ����ʦ <%if rs("job") ="ͨѶ����ʦ" then Response.Write "selected"%>>ͨѶ����ʦ
			  <OPTION value=��Ƭ��/DSP/�ײ�������� <%if rs("job") ="��Ƭ��/DSP/�ײ��������" then Response.Write "selected"%>>��Ƭ��/DSP/�ײ��������
			  <OPTION>-----------��еרҵ��Ա-----------
			  <OPTION value=��е����ʦ <%if rs("job") ="��е����ʦ" then Response.Write "selected"%>>��е����ʦ
			  <OPTION value=ģ�߹���ʦ <%if rs("job") ="ģ�߹���ʦ" then Response.Write "selected"%>>ģ�߹���ʦ
			  <OPTION value=���繤��ʦ <%if rs("job") ="���繤��ʦ" then Response.Write "selected"%>>���繤��ʦ
			  <OPTION value=���ֳ���/��������� <%if rs("job") ="���ֳ���/���������" then Response.Write "selected"%>>���ֳ���/���������
			  <OPTION>-------���ز�/����רҵ��Ա-------
			  <OPTION value=���ز�����/�߻� <%if rs("job") ="���ز�����/�߻�" then Response.Write "selected"%>>���ز�����/�߻�
			  <OPTION value=���ز�����/���� <%if rs("job") ="���ز�����/����" then Response.Write "selected"%>>���ز�����/����
			  <OPTION value=����/�ṹ����ʦ <%if rs("job") ="����/�ṹ����ʦ" then Response.Write "selected"%>>����/�ṹ����ʦ
			  <OPTION value=���̼���ʦ <%if rs("job") ="���̼���ʦ" then Response.Write "selected"%>>���̼���ʦ
			  <OPTION value=����Ԥ���� <%if rs("job") ="����Ԥ����" then Response.Write "selected"%>>����Ԥ����
			  <OPTION value=����ˮ/ˮ�繤��ʦ <%if rs("job") ="����ˮ/ˮ�繤��ʦ" then Response.Write "selected"%>>����ˮ/ˮ�繤��ʦ
			  <OPTION value=����ůͨ <%if rs("job") ="����ůͨ" then Response.Write "selected"%>>����ůͨ
			  <OPTION value=��ҵ���� <%if rs("job") ="��ҵ����" then Response.Write "selected"%>>��ҵ����
			  <OPTION>--------����/����רҵ��Ա--------
			  <OPTION value=����ҵ <%if rs("job") ="����ҵ" then Response.Write "selected"%>>����ҵ
			  <OPTION value=֤ȯ�ڻ� <%if rs("job") ="֤ȯ�ڻ�" then Response.Write "selected"%>>֤ȯ�ڻ�
			  <OPTION value=����ҵ <%if rs("job") ="����ҵ" then Response.Write "selected"%>>����ҵ
			  <OPTION value=˰����Ա <%if rs("job") ="˰����Ա" then Response.Write "selected"%>>˰����Ա
			  <OPTION value=��������/������Ա <%if rs("job") ="��������/������Ա" then Response.Write "selected"%>>��������/������Ա
			  <OPTION>------�Ľ�����/����רҵ��Ա------
			  <OPTION value=����/���� <%if rs("job") ="����/����" then Response.Write "selected"%>>����/����
			  <OPTION value=�㲥����/�Ļ����� <%if rs("job") ="�㲥����/�Ļ�����" then Response.Write "selected"%>>�㲥����/�Ļ�����
			  <OPTION value=�ߵȽ��� <%if rs("job") ="�ߵȽ���" then Response.Write "selected"%>>�ߵȽ���
			  <OPTION value=�еȽ��� <%if rs("job") ="�еȽ���" then Response.Write "selected"%>>�еȽ���
			  <OPTION value=Сѧ/�׶����� <%if rs("job") ="Сѧ/�׶�����" then Response.Write "selected"%>>Сѧ/�׶�����
			  <OPTION value=ְҵ����/��ѵ/�ҽ� <%if rs("job") ="ְҵ����/��ѵ/�ҽ�" then Response.Write "selected"%>>ְҵ����/��ѵ/�ҽ�
			  <OPTION value=������ <%if rs("job") ="������" then Response.Write "selected"%>>������
			  <OPTION value=����ҽ�� <%if rs("job") ="����ҽ��" then Response.Write "selected"%>>����ҽ��
			  <OPTION value=��ʦ/���ɹ��� <%if rs("job") ="��ʦ/���ɹ���" then Response.Write "selected"%>>��ʦ/���ɹ���
			  <OPTION>----------����ҵרҵ��Ա----------
			  <OPTION value=����/���� <%if rs("job") ="����/����" then Response.Write "selected"%>>����/����
			  <OPTION value=�Ƶ�/���� <%if rs("job") ="�Ƶ�/����" then Response.Write "selected"%>>�Ƶ�/����
			  <OPTION value=Ѱ��/��Ѷ <%if rs("job") ="Ѱ��/��Ѷ" then Response.Write "selected"%>>Ѱ��/��Ѷ
			  <OPTION value=����������ҵ <%if rs("job") ="����������ҵ" then Response.Write "selected"%>>����������ҵ
			  <OPTION>----------����רҵ��Ա ----------
			  <OPTION value=����/��Դ <%if rs("job") ="����/��Դ" then Response.Write "selected"%>>����/��Դ
			  <OPTION value=����ѧ���� <%if rs("job") ="����ѧ����" then Response.Write "selected"%>>����ѧ����
			  <OPTION value=�������� <%if rs("job") ="��������" then Response.Write "selected"%>>��������
			  <OPTION value=������ҩ <%if rs("job") ="������ҩ" then Response.Write "selected"%>>������ҩ
			  <OPTION value=��漼�� <%if rs("job") ="��漼��" then Response.Write "selected"%>>��漼��
			  <OPTION value=���ż��� <%if rs("job") ="" then Response.Write "selected"%>>���ż���
			  <OPTION value=����/���й滮 <%if rs("job") ="����/���й滮" then Response.Write "selected"%>>����/���й滮
			  <OPTION value=����/��� <%if rs("job") ="����/���" then Response.Write "selected"%>>����/���
			  <OPTION value=��ʳ/ʳƷ/�Ǿ� <%if rs("job") ="��ʳ/ʳƷ/�Ǿ�" then Response.Write "selected"%>>��ʳ/ʳƷ/�Ǿ�
			  <OPTION value=��֯��װ���� <%if rs("job") ="��֯��װ����" then Response.Write "selected"%>>��֯��װ����
			  <OPTION value=��װ/ӡˢ/��ֽ <%if rs("job") ="��װ/ӡˢ/��ֽ" then Response.Write "selected"%>>��װ/ӡˢ/��ֽ
			  <OPTION value=ұ��/��Ϳ/�������� <%if rs("job") ="ұ��/��Ϳ/��������" then Response.Write "selected"%>>ұ��/��Ϳ/��������
			  <OPTION value=��ȫ���� <%if rs("job") ="��ȫ����" then Response.Write "selected"%>>��ȫ����
			  <OPTION value=ũ������/԰��/԰�� <%if rs("job") ="ũ������/԰��/԰��" then Response.Write "selected"%>>ũ������/԰��/԰��
			  <OPTION value=��ͨ���䣨��½�գ� <%if rs("job") ="��ͨ���䣨��½�գ�" then Response.Write "selected"%>>��ͨ���䣨��½�գ�</select></p>
              <p align="center">��Ƹ����:<INPUT maxLength=3    
                  size=3 name=zpnum value="<%=rs("zpnum")%>">�� �����ص�:<B><SELECT size=1 name=gzdd>               
              <OPTION>���������б���ѡ��</OPTION>
			  <OPTION value=������ <%if rs("gzdd") ="������" then Response.Write "selected"%>>������
              <OPTION value=����� <%if rs("gzdd") ="�����" then Response.Write "selected"%>>�����
              <OPTION value=�Ϻ��� <%if rs("gzdd") ="�Ϻ���" then Response.Write "selected"%>>�Ϻ���
			  <OPTION value=������ <%if rs("gzdd") ="������" then Response.Write "selected"%>>������
              <OPTION value=����ʡ <%if rs("gzdd") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=����ʡ <%if rs("gzdd") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=����ʡ <%if rs("gzdd") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=�㶫ʡ <%if rs("gzdd") ="�㶫ʡ" then Response.Write "selected"%>>�㶫ʡ
			  <OPTION value=�Ĵ�ʡ <%if rs("gzdd") ="�Ĵ�ʡ" then Response.Write "selected"%>>�Ĵ�ʡ
              <OPTION value=����ʡ <%if rs("gzdd") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=����ʡ <%if rs("gzdd") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=����ʡ <%if rs("gzdd") ="����ʡ" then Response.Write "selected"%>>����ʡ
			  <OPTION value=����ʡ <%if rs("gzdd") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=����ʡ <%if rs("gzdd") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=����ʡ <%if rs("gzdd") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=����ʡ <%if rs("gzdd") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=�ຣʡ <%if rs("gzdd") ="�ຣʡ" then Response.Write "selected"%>>�ຣʡ
			  <OPTION value=ɽ��ʡ <%if rs("gzdd") ="ɽ��ʡ" then Response.Write "selected"%>>ɽ��ʡ
              <OPTION value=ɽ��ʡ <%if rs("gzdd") ="ɽ��ʡ" then Response.Write "selected"%>>ɽ��ʡ
			  <OPTION value=����ʡ <%if rs("gzdd") ="����ʡ" then Response.Write "selected"%>>����ʡ
              <OPTION value=����ʡ <%if rs("gzdd") ="����ʡ" then Response.Write "selected"%>>����ʡ
			  <OPTION value=�㽭ʡ <%if rs("gzdd") ="�㽭ʡ" then Response.Write "selected"%>>�㽭ʡ
			  <OPTION value=������ʡ <%if rs("gzdd") ="������ʡ" then Response.Write "selected"%>>������ʡ
			
             </SELECT></B>
              </p>
              <p align="center">��λ����:<br>
             <%if rs("zptext")<>"" then
			 kzptext=replace(rs("zptext"),"<br>",chr(13))
             kzptext=replace(kzptext,"&nbsp;"," ")
			 else
			 kzptext="" %>
             <%end if%>      
			 <textarea rows="6" name="zptext" cols="41"><%=kzptext%></textarea>      
             </p>
             <p align="center">���Ҫ��:<br>
             <%if rs("xgyq")<>"" then
	         kxgyq=replace(rs("xgyq"),"<br>",chr(13))
             kxgyq=replace(kxgyq,"&nbsp;"," ")
			 else
			 kxgyq=""%>
             <%end if%>      
			 <textarea rows="6" name="xgyq" cols="41"><%=kxgyq%></textarea>      
              </p>
              <p align="center"><input type="button" value="�� ��" onClick="check()">

              <p align="center">��</td>
          </tr>
        </table>
        </div>
        <p>��
      </td>
  <center>
      <td width="1" height="351" valign="top" bgcolor="#00006A"></td>
      <td width="124" height="351" valign="top" bgcolor="#F3F3F3">��</td>
    </tr>
    <tr>
      <td width="780" height="1" valign="top" colspan="5" bgcolor="#53BEB0"> 
        <p align="center"> </td>
    </tr>
    <tr>
      <td width="780" height="3" valign="top" colspan="5">
      <p align="center">
      </td>
    </tr>
    <tr>
      <td width="780" height="7" valign="top" colspan="5">
      <p align="center"><script language="javascript" src="../inc/copyright.js"></script>
      </td>
    </tr>
    <tr>
      <td width="780" height="3" valign="top" colspan="5">
      </td>
    </tr>
  </table>
  </center>
</div>
</form>
</body>

</html>

<% zptext=htmlencode2(request("zptext"))
if zptext="" then Response.End
   job=request("job")
   zpnum=request("zpnum")
   gzdd=request("gzdd")
   xgyq=htmlencode2(request("xgyq"))

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from company where uname='"&uname&"'"
rs.open sql,conn,3,3
rs("zptext")=zptext
rs("job")=job
rs("zpnum")=zpnum
rs("gzdd")=gzdd
rs("xgyq")=xgyq
rs("idate")=date()
rs.update
rs.close
Response.Redirect "main.asp"
%>
