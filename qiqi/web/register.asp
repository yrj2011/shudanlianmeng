<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="inc/index_conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="config.asp"-->
<!--#include file="md5.asp"-->
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷�����  �� �� ý
'�ٷ���ҳ��http://www.778892.com
'����֧�֣�xsh
'QQ;859680966
'�绰��15889679112
'��Ȩ����Ȩ�������ߴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2013 ���ߴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
promotion=request.QueryString("promotion")
if promotion<>"" then
Sql = "select username from "&jieducm&"_register where id="&promotion&""
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF ( Rs.Eof or Rs.Bof) Then
response.write "<script>alert('�޴��Ƽ��ˣ�');window.location.href='register.asp';</script>"  
response.End()
else
session("tjrname")=rs("username")
end if				
end if

zhuce=request.Form("zhuce")
if zhuce="ok" then
 session("jieducm_ok")="jieducm_reg"
 call register()    
end if
 %>
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<TITLE><%=stupname%>-<%=stuptitle%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK rev=Stylesheet href="img/global.css" type=text/css rel=Stylesheet>
<LINK href="css/login.css" type=text/css rel=stylesheet>
<LINK href="css/top_bottom.css" type=text/css rel=stylesheet>
<LINK href="css/layout.css" type=text/css rel=stylesheet>
<LINK rel=stylesheet type=text/css href="css/global.css"> 
<SCRIPT language=JavaScript src="js/jquery.min.js"></SCRIPT>
<SCRIPT src="js/jieducm_pupu.js" type=text/javascript></SCRIPT>
<script language="JavaScript" type="text/JavaScript" src="js/js.js"></script>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR></HEAD>
<BODY  onload=document.form.UserID.focus()> 

<SCRIPT language=javascript>
function xiechuemail(){
document.getElementById("Email").value=document.getElementById("QQ").value+'@qq.com';
}
function  save_onclick()
{
  var strTemp = document.form.UserID.value;
  if (strTemp.length == 0 )
  {
      alert("�������û�����");
      document.form.UserID.focus();
      return false;
  }

if(form.UserID.value.replace(" ","").replace( ).length!=form.UserID.value.length)
{ alert('�û��������пո�');
  document.form.UserID.focus();
      return false;
}
			
   if (strTemp.length < 3 || strTemp.length>16 )
  {
      alert("�û����ĳ��ȱ���Ϊ3��16���ַ���һ�������������ַ���");
      document.form.UserID.focus();
      return false;
  }
  var strTemp = document.form.password.value;
  if (strTemp.length == 0 )
  {
      alert("���������룡");
      document.form.password.focus();
      return false;
  }
  
    if (strTemp.length < 6 || strTemp.length>16 )
  {
      alert("���������6��16λ���ֻ���ĸ��");
      document.form.password.focus();
      return false;
  }
  
  var strTemp1 =document.form.password2.value;
  if (strTemp1.length == 0 )
  {
      alert("�������ظ��������룡");
      document.form.password2.focus();
      return false;
  }
  
  if (strTemp!=strTemp1)
  {
      alert("�����������벻һ�£�");
      document.form.password.focus();
      return false;
  }
  
  var qq=document.form.QQ.value
  if(qq.length==0)
  {
     alert("QQ����Ϊ��!");
	 document.form.QQ.focus();
	 return false;
	 }
	   var Email=document.form.Email.value
  if(Email.length==0)
  {
     alert("Email����Ϊ��!");
	 document.form.Email.focus();
	 return false;
	 }
	 
	  if(document.form.Email.value.length!=0)
    if (document.form.Email.value.charAt(0)=="." ||        
         document.form.Email.value.charAt(0)=="@"||       
         document.form.Email.value.indexOf('@', 0) == -1 || 
         document.form.Email.value.indexOf('.', 0) == -1 || 
         document.form.Email.value.lastIndexOf("@")==document.form.Email.value.length-1 || 
         document.form.Email.value.lastIndexOf(".")==document.form.Email.value.length-1)
     {
      alert("Email��ַ��ʽ����ȷ��");
      document.form.Email.focus();
      return false;
      }
	  
    var czm =document.form.czm.value;
  if (czm.length == 0)
  {
      alert("�����벻��Ϊ�գ�");
      document.form.czm.focus();
      return false;
  }
if (czm.length < 6 || czm.length>16 )
  {
      alert("�����������6��16λ���ֻ���ĸ��");
      document.form.czm.focus();
      return false;
  }
   var czm2 =document.form.czm2.value;
  if (czm2.length == 0)
  {
      alert("ȷ�ϲ���������벻��Ϊ�գ�");
      document.form.czm2.focus();
      return false;
  }
  if (czm!=czm2)
  {
      alert("���β��������벻һ�£�");
      document.form.czm.focus();
      return false;
  }
  
  if (strTemp==czm)
  {
      alert("�����벻����͵�¼������ͬ��");
      document.form.czm.focus();
      return false;
  }
  var phone =document.form.phone.value;
  if (phone.length == 0)
  {
      alert("�ֻ�����Ϊ�գ�");
      document.form.phone.focus();
      return false;
  }
  if (phone.length !=11 )
  {
      alert("����д��ȷ���ֻ����룡");
      document.form.phone.focus();
      return false;
  }
  
<%if stupCheckCode="2" then%>
if (document.form.CheckCode)
	{
		if (document.form.CheckCode.value==""){
		   alert ("������������֤�룡");
		   document.form.CheckCode.focus();
		   return(false);
		}
	}
<%end if%>
  return true;  
}

	function SwitchDivDisplay(divName) {
		var obj_reg_info = document.getElementById(divName);
		//alert(obj_reg_info);
	   	if(obj_reg_info.style.display=="none") {
	       	obj_reg_info.style.display="inline";
	   	} else {
	   		obj_reg_info.style.display="block";
	   		obj_reg_info.style.display="none";
		}
	}

</SCRIPT>
<!--#include file=jieducm_top.asp-->
<TABLE cellSpacing=0 cellPadding=0 width=960 align=center border=0>
  <TBODY>
  <TR>
    <TD width=20 height=32><IMG src="images/Top_10.gif"></TD>
    <TD align=right width=21 background=images/Top_11.gif><IMG height=32 
      src="images/Top_9.gif" width=10></TD>
    <TD class=K_mttitle align=middle width=120 
      background=images/Top_12.gif>��Աע��</TD>
    <TD align=left width=47 background=images/Top_11.gif><IMG height=32 
      src="images/Top_13.gif" width=10></TD>
    <TD align=left width=728 background=images/Top_11.gif></TD>
    <TD align=right width=24 background=images/Top_11.gif><IMG height=32 
      src="images/Top_14.gif" width=6></TD></TR>
  <TR>
    <TD class=K_mtcontent align=left colSpan=6> 
	   <%if  stupzhu="2" then%>

<DIV class=register>
<H2>��Աע��<SPAN class=current>1.��д��Ա��Ϣ</SPAN><SPAN>2.ͨ���ʼ�ȷ��</SPAN><SPAN>3.ע��ɹ�</SPAN></H2>
<FORM id=form name=form onSubmit="return save_onclick()" method=post>
<INPUT type=hidden value=ok name=zhuce> 
<UL><li style="line-height:22px;"><span class="font14b2">���û����ע��-����ע��10���Ӹ㶨</span><br />
          ����д������Ϣ��<span class="red-bcolor">*</span>Ϊ�������ݡ�<br />
          ��������ϸ����д������Ϣ����ʵ�ĸ�����Ϣ�����ڸ���ʹ�õķ����������ı����Լ���ݣ�&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		  <a href="/news.asp?/1465.html" target="_blank"><strong>��ʾ �Ƿ��ַ����ͻ�����</strong></a></li> 
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=username>��Ա����</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=UserID onblur=name_chk(this) 
  maxLength=12 name=UserID>
  <SCRIPT id=tname></SCRIPT>
  
  <DIV class=form-state id=c0><FONT color=#ff0000>*</FONT> 4-12���ַ�</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=password>��¼���룺</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=password type=password 
  name=password> 
  <DIV class=form-state><FONT color=#ff0000>*</FONT> ��¼ʱ��Ҫʹ������</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL 
  for=verify-password>ȷ�����룺</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=password2 type=password 
  name=password2> 
  <DIV class=form-state><FONT color=#ff0000>*</FONT> �ظ����������</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=password2>�������룺</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=czm name=czm type=password > 
  <DIV class=form-state><FONT color=#ff0000>*</FONT> �������޸�����Ȳ�����Ҫ�õ�</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=password2>ȷ�ϲ����룺</LABEL> 

  <DIV class=form-entry><INPUT class=txt id=czm2 name=czm2 type=password > 
  <DIV class=form-state><FONT color=#ff0000>*</FONT> �ظ�����Ĳ�����</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=qq>QQ���룺</LABEL> 
  <DIV class=form-entry>
  <input   class=txt name="QQ" type="text" id="QQ" size="30" onKeyUp="if(isNaN(value))execCommand('undo')"  onBlur="xiechuemail()"/>
  <DIV class=form-state><FONT color=#ff0000>*</FONT> QQ���벻��Ϊ�գ�˫����ϵ�ر���</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=email>�������䣺</LABEL> 
  <DIV class=form-entry><input class=txt name="Email" type="text" id="Email" size="30" />
  <DIV class=form-state><FONT color=#ff0000>*</FONT> �����˺ź�ȡ������ʱ��Ҫ�õ��������Ƿ���ȷ��</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=tjr>�Ƽ��ˣ�</LABEL> 
  <DIV class=form-entry><input class=txt name="tjr" type="text" id="tjr" value="<%=session("tjrname")%>" size="30"  <%if session("tjrname")<>"" then%> readonly <%end if%>  />
  <DIV 
  class=form-state>û�п����գ����ڱ��Ƽ��ˣ�û��������ʧ���Ƽ��˵õ����ֽ���</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' onmouseout='this.setAttribute("class","")'><LABEL for=phone>�ֻ����룺</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=phone  onkeyup="if(isNaN(value))execCommand('undo')" name=phone > 
  <DIV class=form-state>��ƾ�ֻ��һ�����</DIV></DIV></LI>
  
  <LI class=random-code><LABEL for=username>��֤�룺</LABEL> 
  <DIV class=form-entry><INPUT class=txt maxLength=4 id=CheckCode onKeyUp="if(isNaN(value))execCommand('undo')"  name=CheckCode> 
  <img src="jieducm_code.asp" alt="��֤��" onClick="this.src='jieducm_code.asp?rnd=' + Math.random();" /> 
  <SPAN>�����壿���ͼƬ��һ��</SPAN> </DIV></LI>
  
  <LI class=random-code><LABEL for=checkbox>��д���ࣺ</LABEL> 
  <DIV class=form-entry><INPUT id=checkbox onClick="javascript:SwitchDivDisplay('showinfo_c1')" type=checkbox name=checkbox> ���������Ϣ��ѡ� </DIV></LI></UL>
<UL id=showinfo_c1 style="DISPLAY: none">
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=xt><SPAN 
  class=font14b2>ѡ����Ŀ��</SPAN></LABEL> 
  <DIV class=form-entry>
  <HR style="COLOR: #ff9933">
  </DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=weiti>�������⣺</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=weiti name=weiti>
  <DIV class=form-state>�����������ʾ����</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=daai>����𰸣�</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=daai name=daai> 
  <DIV class=form-state>��������������</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=rname>��ʵ������</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=rname name=rname> 
  <DIV class=form-state><FONT color=#ff0000>*</FONT> 
  ��ʵ����</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=HomePage>���̵�ַ��</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=HomePage value=http:// 
  name=HomePage> 
  <DIV class=form-state><FONT color=#ff0000>*</FONT> 
  ���̵�ַ</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=shopnote>����������</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=shopnote name=shopnote>
  <DIV class=form-state><FONT color=#ff0000>*</FONT> 
  ��������</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=sex>�Ա�</LABEL> 
  <DIV class=form-entry><INPUT id=radio type=radio CHECKED value=1 name=sex> ��ʿ 
  <INPUT id=radio2 type=radio value=2 name=sex> Ůʿ </DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=address>��ϵ��ַ��</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=address name=address>
  <DIV class=form-state><FONT color=#ff0000>*</FONT> 
  ��ϵ��ַ</DIV></DIV></LI></UL>
  <div align="center"><a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=stupqq%>&site=qq&menu=yes">��ϵQQ��<%=stupqq%>  ��֤ ��ͨ����</a></div>
<UL>
  <LI>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <INPUT class=submit style="MARGIN-LEFT: 20px; CURSOR: pointer" type="submit" value=ȷ��ע��>
  <INPUT class=submit style="MARGIN-LEFT: 20px; CURSOR: pointer" onClick="window.showModelessDialog('images/jieducm_zhidu.htm',window,'scroll:yes;resizable:yes;help:no;dialogWidth:650px;dialogHeight:450px')" type=button value=ע��Э��>&nbsp;&nbsp;
  </LI>
</UL></FORM></DIV>
	<%elseif  stupzhu="1" then%>
    <table width="900" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="100" colspan="3"><div align="center"><span class="font14b2"> ϵͳ��ͣע��</span><br />
        </div></td>
        </tr>
        </table>
   
    <%end if%>
	</TD></TR></TBODY></TABLE><br>
<!--#include file=jieducm_bottom.asp-->
<%
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>
