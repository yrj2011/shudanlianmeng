<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->
<!--#include file="../md5.asp"-->
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
if request.Form<>"" then
login_ip=HtmlEncode(Replace(Trim(Request.Form("login_ip")),"'","''"))
if czm<>request.form("czm") then
	Response.Write("<script>alert('�����벻��ȷ!');history.back();</script>")
    response.End()
end if

if  checkenter(pwd) then
else
	response.write "<script language=javascript>alert('�벻Ҫ����Ƿ��ַ���');history.back();</script>"
	response.end
end if	

if  checkenter(eczm) then
else
	response.write "<script language=javascript>alert('�벻Ҫ����Ƿ��ַ���');history.back();</script>"
	response.end
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select login_ip From "&jieducm&"_register where username='" &session("useridname")& "'" ,Conn,3,3 
rs("login_ip")=login_ip
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write("<script>alert('�޸ĳɹ�!');window.location.href='setip.asp';</script>")
response.End()
end if
%>	

<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="../css/center.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   
<!--#include file=../inc/jieducm_top.asp-->
<SCRIPT language=javascript>

function  save_onclick()
{
var login_ip=document.form.login_ip.value
var czm=document.form.czm.value
  
if (login_ip.length==0 )
   {
     alert("�������½IP!");
	 document.form.login_ip.focus();
	 return false;
     }

if (czm.length==0 )
  {
      alert("����������룡");
      document.form.czm.focus();
      return false;
  }	
   
  return true;  
}

</SCRIPT>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<tr>
	    <td>
        <table width="910"  border="0" cellspacing="0" cellpadding="0">
	          <tr>
	          
	            <td width="140" valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"> </td>
    </tr>
</table>



<table id="menuToolBar" width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
    <td valign="top" bgcolor="FFFFFF">
    <!--#include file="userleft.asp"--></TD>
</TR>
</table></td>
	            <td width="15">&nbsp;</td>
	          
	            <td valign="top">
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"></td>
    </tr>
</table>
                <div class=pp9>
                  <div style="PADDING-BOTTOM: 15px; WIDTH: 97%">
                    <div class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt; ���õ�½IP &gt;&gt; </div>
                    <div class=pp8><strong>���õ�½IP</strong></div>


<FORM name="form" id="form" onSubmit="return save_onclick()" method=post action="">
<table width="99%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="10" colspan="3">&nbsp;</td>
        </tr>
      <tr>
        <td width="146" height="40" align="right" class="font12h">��½IP��</td>
       <td width="230">
	   <input name="login_ip" type="text" class="text_normal" id="login_ip" size="30"  value="<%=login_ip%>"/>	   </td>
        <td width="524">�㱾�ε�½��IP�ǣ�<%=Request.ServerVariables("REMOTE_ADDR")%>&nbsp; 0��ʾû�����õ�½ip</td>
      </tr>
      <tr>
        <td height="40" align="right" class="font12h">�����룺</td>
        <td><input name="czm" type="password" id="czm" size="30"  /></td>
        <td>&nbsp;</td>
      </tr>
      
      <tr>
        <td height="40" align="right" class="font12h">&nbsp;</td>
        <td align="center" class="red-bcolor">
		<input style="FONT-WEIGHT: bold; WIDTH: 120px; CURSOR: pointer; COLOR: #000000; HEIGHT: 26px" type="submit" name="button" id="button"  value="ȷ��">		</td>
        <td>&nbsp;</td>
      </tr>
    </table>
					</FORM>
                  </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
call closeconn()%>
</BODY></HTML>
