<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="inc/index_conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="config.asp"-->
<!--#include file="md5.asp"-->
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ**************************************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷����ݶȴ�ý
'�ٷ���ҳ��http://www.jieducm.cn
'����֧�֣�robin 
'QQ;616591415
'�绰��13598006137
'��Ȩ����Ȩ���ڽݶȴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2008��2011 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------	
action=request.Form("action")
if action="ok" then
call login()
end if
 %>
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<TITLE><%=stupname%>-<%=stuptitle%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="css/Css.css" type=text/css rel=stylesheet>
<LINK href="css/login.css" type=text/css rel=stylesheet>
<LINK href="css/top_bottom.css" type=text/css rel=stylesheet>
<LINK rel=stylesheet type=text/css href="css/global.css"> 
<SCRIPT language=JavaScript src="js/jquery.min.js"></SCRIPT>
<SCRIPT src="js/jieducm_pupu.js" type=text/javascript></SCRIPT>
</HEAD>
<BODY  onload=document.form.name.focus()> 
<script language=javascript>
<!--

function CheckForm()
{
	if(document.form.name.value=="")
	{
		alert("�������û�����");
		document.form.name.focus();
		return false;
	}
	if(document.form.pass.value == "")
	{
		alert("���������룡");
		document.form.pass.focus();
		return false;
	}
	if (document.form.CheckCode)
	{
		if (document.form.CheckCode.value==""){
		   alert ("������������֤�룡");
		   document.form.CheckCode.focus();
		   return(false);
		}
	}
}
//-->
</script>
<!--#include file=jieducm_top.asp-->
<TABLE cellSpacing=0 cellPadding=0 width=960 align=center border=0>
  <TBODY>
  <TR>
    <TD width=20 height=32><IMG src="images/Top_10.gif"></TD>
    <TD align=right width=21 background=images/Top_11.gif><IMG height=32 
      src="images/Top_9.gif" width=10></TD>
    <TD class=K_mttitle align=middle width=120 
      background=images/Top_12.gif>��Ա��½</TD>
    <TD align=left width=47 background=images/Top_11.gif><IMG height=32 
      src="images/Top_13.gif" width=10></TD>
    <TD align=left width=728 background=images/Top_11.gif></TD>
    <TD align=right width=24 background=images/Top_11.gif><IMG height=32 
      src="images/Top_14.gif" width=6></TD></TR>
  <TR>
    <TD class=K_mtcontent align=left colSpan=6>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD align=middle width="48%" height=350>
            <FORM name="form" method="post" onSubmit="return CheckForm();"  action="">
			<input name="action" type="hidden" value="ok">
			<table width="90%" border="0" align="right" cellpadding="8" cellspacing="0" >
              <tr>
                <td colspan="2" bgcolor="#FFFFFF"><div align="left"><img src="img/login.jpg" width="270" height="57"></div></td>
                </tr>
              <tr>
                <td width="25%" bgcolor="#FFFFFF">��Ա�˺ţ�</td>
                <td width="75%" bgcolor="#FFFFFF"><div align="left"><tt>
                  <input name=name  class=username type=text size="30" maxlength=16 style="WIDTH: 200px;" >
                </tt></div></td>
              </tr>
              <tr>
                <td bgcolor="#FFFFFF">��¼���룺</td>
                <td bgcolor="#FFFFFF"><div align="left"><tt>
                  <input name=pass class=password type=password size="30" maxlength=16 style="WIDTH: 200px;">
                </tt></div></td>
              </tr>
              <tr>
                <td bgcolor="#FFFFFF">��֤�룺</td>
                <td bgcolor="#FFFFFF"><div align="left"><tt>
                  <input  type=text maxlength=4 name=CheckCode class=code onKeyUp="if(isNaN(value))execCommand('undo')">
                </tt><img src="jieducm_code.asp" alt="��֤��" onClick="this.src='jieducm_code.asp?rnd=' + Math.random();" /></div></td>
              </tr>
              <tr>
                <td colspan="2" bgcolor="#FFFFFF"><label>
                <a onClick="javascript:window.open('getpwd.asp','shouchang','width=450,height=300');" href="#"><img src="images/forgetpw.gif" width="12" height="13" border="0">&nbsp;�һ�����</a>&nbsp;<a href="register.asp">&nbsp;&nbsp;&nbsp;<img src="images/register.gif" width="14" height="13" border="0">&nbsp;����ע��</a></label></td>
              </tr>
              <tr>
                <td colspan="2" bgcolor="#FFFFFF"><INPUT style="FONT-WEIGHT: bold; WIDTH: 120px; CURSOR: pointer; COLOR: #000000; HEIGHT: 26px"  type=submit value="��½����ƽ̨  " name=ctl00$AllBaseContentPlaceHolder$ctl00>
                  &nbsp;
                  &nbsp;
                   <a href="register.asp"><INPUT style="WIDTH: 100px; CURSOR: pointer; HEIGHT: 26px" type=reset value=" �� �� �� �� "></a></td>
              </tr>
            </table>
			</FORM></TD>
          <TD width="1%">
            <DIV 
            style="WIDTH: 2px; HEIGHT: 230px; BACKGROUND-COLOR: #3da3e7"></DIV></TD>
          <TD class=riglogin width="51%"><B>ƽ̨��Աͬʱ���ܾ��ʷ���</B><BR>
            <IMG 
            src="images/jieducm_b0.gif"> <TT> ��һ����</STRONG><A href="<%=stupurl%>/register.asp">ע��</A>Ϊ��վ��Ա����½ϵͳ��</TT><BR>
              <IMG 
            src="images/jieducm_b1.gif"> <TT>�ڶ�������ˢ�����Ľ�����,��һ������ϵͳ���Զ�Ϊ�����ӷ����㡣</TT><BR>
              <IMG 
            src="images/jieducm_b2.gif"> <TT>&nbsp;������������ղŹ���ı�������"����"</TT><BR><IMG 
            src="images/jieducm_b3.gif"> <TT>���Ĳ���ӵ�з�����󼴿ɷ�������ƽ̨������Ϊ�û��Զ���</TT> 
      </TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE><br>
<!--#include file=jieducm_bottom.asp-->
<%
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>
