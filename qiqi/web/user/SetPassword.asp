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
pwd=HtmlEncode(trim(request.form("pwd")))
eczm=HtmlEncode(trim(request.form("eczm")))
if czm<>request.form("czm") then
	Response.Write("<script>alert('��������ȷ�ĵ�ǰ������!');history.back();</script>")
    response.End()
elseif pwd=eczm then
	Response.Write("<script>alert('�����벻����͵�¼������ͬ!');history.back();</script>")
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
	    rs.open "Select * From "&jieducm&"_register where username='" &session("useridname")& "'" ,Conn,3,3 
		if pwd<>"" then 
    	rs("password1")= md5(pwd)
		end if
		if eczm<>"" then
		rs("czm")=eczm
		end if
    	rs.update
    	rs.close
		Response.Write("<script>alert('�޸ĳɹ�!');window.location.href='setpassword.asp';</script>")
		response.End()
set rs=nothing
conn.close
set conn=nothing
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
var czm=document.form.czm.value
var pwd=document.form.pwd.value
var pwd2=document.form.pwd2.value
var eczm=document.form.eczm.value
var eczm2=document.form.eczm2.value

if(czm.length==0)
  {
     alert("��ǰ�����벻��Ϊ��!");
	 document.form.czm.focus();
	 return false;
	 }	
	 
if(eczm.length==0 && eczm2.length==0 && pwd.length==0 && pwd2.length==0)
  {
     alert("Ҫ�޸ĵĵ�¼����Ͳ�������ͬʱΪ��!");
	 document.form.pwd.focus();
	 return false;
	 }	
if (pwd.length!=0 && pwd.length<6 )
   {
     alert("��¼���������6��16λ���ֻ���ĸ!");
	 document.form.pwd.focus();
	 return false;
     }
if (pwd.length!=0 && pwd2.length==0 )
   {
     alert("������ȷ�ϵ�¼����!");
	 document.form.pwd2.focus();
	 return false;
     }

if (pwd.length!=0 && pwd!=pwd2)
  {
      alert("�����������벻һ�£�");
      document.form.pwd.focus();
      return false;
  } 
  
if (eczm.length!=0 && eczm.length<6 )
   {
     alert("�����������6��16λ���ֻ���ĸ!");
	 document.form.eczm.focus();
	 return false;
     }
if (eczm.length!=0 && eczm2.length==0 )
   {
     alert("������ȷ�ϲ�������!");
	 document.form.eczm2.focus();
	 return false;
     }

if (eczm.length!=0 && eczm!=eczm2)
  {
      alert("���β��������벻һ�£�");
      document.form.eczm.focus();
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
                    <div class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt; �޸����� &gt;&gt; </div>
                    <div class=pp8><strong>�޸�����</strong></div>


<FORM name="form" id="form" onSubmit="return save_onclick()" method=post action="">
<table width="99%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="10" colspan="3">&nbsp;</td>
        </tr>
      <tr>
        <td width="146" height="40" align="right" class="font12h">��ǰ�����룺</td>
       <td width="230">
	   <input name="czm" type="password" class="text_normal" id="czm" size="30" />
	   </td>
        <td width="524">�޸�������Ҫ�ṩ��ǰ������</td>
      </tr>
      <tr>
        <td height="40" align="right" class="font12h">&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="40" align="right" class="font12h">��¼���룺</td>
        <td><input name="pwd" type="password" id="pwd" size="30"  /></td>
        <td>�������޸�</td>
      </tr>
      
      <tr>
        <td height="40" align="right" class="font12h">ȷ�ϵ�¼���룺</td>
        <td><span class="font12h">
          <input name="pwd2" type="password" id="pwd2" size="30"  />
        </span></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="40" align="right" class="font12h">&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      
  
      <tr>
        <td height="40" align="right" class="font12h">�����룺</td>
        <td align="left" class="red-bcolor"><input name="eczm" type="password" id="eczm" size="30"  /></td>
        <td>�������޸�</td>
      </tr>
      <tr>
        <td height="40" align="right" class="font12h">ȷ�ϲ����룺</td>
        <td align="left" class="red-bcolor"><input name="eczm2" type="password" id="eczm2" size="30"  /></td>
        <td>�����벻����͵�¼������ͬ</td>
      </tr>
      <tr>
        <td height="40" colspan="3" align="left" class="font12h">&nbsp;</td>
      </tr>
      <tr>
        <td height="40" align="right" class="font12h">&nbsp;</td>
        <td align="center" class="red-bcolor"><input type="submit" name="button" id="button" value="��  ��" /></td>
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
