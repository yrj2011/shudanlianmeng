<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">

<!--#include file="inc/conn.asp"-->
<!--#INCLUDE FILE="inc/md5.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
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
%>
<LINK href="../images/Default.css" type=text/css rel=stylesheet>
<LINK href="../images/Bbs.css" type=text/css rel=stylesheet>
<LINK href="../images/style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
<style type="text/css">
<!--
.STYLE1 {font-size: 16px}
-->
</style>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">



<SCRIPT language=javascript>
function  save_onclick()
{
  var strTemp = document.form.password.value;
  var strTemp1 = document.form.password2.value;
  if (strTemp!= strTemp1 )
  {
      alert("������д���벻һ�£�");
      document.form.password.focus();
      return false;
  }
    var qq1=document.form.QQ.value;
  if(qq1=="")
  {
    alert("����������QQ��");
	document.form.QQ.focus();
	return false;
	}
	    var Email1=document.form.Email.value;
  if(Email1=="")
  {
    alert("����������Email��");
	document.form.Email.focus();
	return false;
	}
 
  var strTemp = document.form.faceheight.value;
  if ((strTemp <1) || (strTemp >130))
  {
      alert("ͷ�񳤶�С��1�����130");
      document.form.faceheight.focus();
      return false;
  }
  var strTemp = document.form.facewidth.value;
  if ((strTemp <1) || (strTemp >130))
  {
      alert("ͷ��߶�С��1�����130");
      document.form.facewidth.focus();
      return false;
  }


  
  return true;  
}
</SCRIPT>
<% 
username=request("username")
Set rs=server.createobject("ADODB.RECORDSET")
   rs.open "Select * From "&jieducm&"_register where username='"&username&"'",Conn,3,3
%>
<FORM action="usersave.asp" method="POST" name="form" id="form" onSubmit="return save_onclick()">
<DIV style="WIDTH: 950px">
<DIV class=regheight 
style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 15px; PADDING-TOP: 15px"><BR>
  <BR>
<DIV style="COLOR: red"><SPAN id=ctl00_AllBaseContentPlaceHolder_MsgLabel 
style="COLOR: red"></SPAN></DIV>
<DIV>
<P align="left">�û�����</P>
<div align="left"><INPUT id=username maxLength=50 
name=username value="<%=rs("username")%>"  readonly></div>
</DIV>
<div align="left"><BR>
</div>
<DIV>
<P align="left">��¼���룺</P>
<div align="left">
  <INPUT id=password type=password 
maxLength=16 name="password"> 
  <FONT 
color=#ff0000>*</FONT>���޸�������<BR>
</div>
</DIV>
<DIV>
<P align="left">�ظ����룺</P>
<div align="left">
  <INPUT id=password2 type=password 
maxLength=16 name=password2 > 
  <FONT 
color=#ff0000>*</FONT> �ظ����������</div>
</DIV>
<div align="left"><BR>
</div>
<DIV>
<P align="left">QQ��</P>
<div align="left">
  <INPUT id=QQ maxLength=15 
name=QQ value="<%=rs("qq")%>"> 
  <FONT color=#ff0000>*</FONT> 
  ˢ��ʱ����<SPAN id=ctl00_AllBaseContentPlaceHolder_RequiredFieldValidator3 
style="DISPLAY: none; COLOR: red">* QQ����Ϊ�գ�</SPAN></div>
</DIV>
<div align="left"><BR>
</div>
<DIV>
<P align="left">�������䣺</P>
<div align="left">
  <INPUT id=Email maxLength=50 
name=Email value="<%=rs("mail")%>"> 
  <FONT color=#ff0000>*</FONT> 
  ȡ������ʹ�ã�ע����ܸ���<SPAN id=ctl00_AllBaseContentPlaceHolder_RequiredFieldValidator4 
style="DISPLAY: none; COLOR: red">* �������䲻��Ϊ�գ�</SPAN><SPAN 
id=ctl00_AllBaseContentPlaceHolder_RegularExpressionValidator1 
style="DISPLAY: none; COLOR: red">* ���������ʽ����ȷ��</SPAN></div>
</DIV>
<div align="left"><BR>
</div>

<DIV>
<P align="left">�������⣺</P>
<div align="left">
  <input name="weiti" type="text" id="weiti" size="30" value="<%=rs("weiti")%>" />
  <FONT color=#ff0000>*</FONT> 
  �����������ʾ����</div>
</DIV>

<div align="left"><BR>
</div>

<DIV>
<P align="left">����𰸣�</P>
<div align="left">
  <input name="daai" type="text" id="daai" size="30"  value="<%=rs("daai")%>"/>
  <FONT color=#ff0000>*</FONT> 
  ��������������</div>
</DIV>

<div align="left"><BR>
</div>

<DIV>
<P align="left">�����룺</P>
<div align="left">
  <INPUT id=czm maxLength=50 
name=czm value="<%=rs("czm")%>">
  <FONT color=#ff0000>*</FONT> 
  ��ˢʱ���� 6-16λ* �����벻��Ϊ�գ�</div>
</DIV>

<BR>
<BR>
<DIV style="COLOR: red">����Ϊѡ��</DIV><BR>
<BR>
<DIV>
<P>�ֻ���</P><INPUT id=phone maxLength=20 
name=phone value="<%=rs("phone")%>"> ����д��ʵ</DIV><BR>
<DIV>
<P>�󶨵��ֻ���</P><INPUT id=dst maxLength=20 
name=dst value="<%=rs("dst")%>"> ����д��ʵ 0��ʾû�а��ֻ�</DIV><BR>
<DIV>
<P>�Ƽ��ˣ�</P><INPUT id=tjr maxLength=20 
name=tjr value="<%=rs("tjr")%>"> ����д��ʵ</DIV>
<br>
<DIV>
<P>��ʵ������</P><INPUT id=rname maxLength=20 
name=rname value="<%=rs("rname")%>"> ����д��ʵ</DIV><BR>
<DIV>
<P>���̵�ַ��</P><input type="text" name="HomePage" size="30" Value="<%=rs("homepage")%>" > ��д�Ļ������ǿ��԰����ƹ��</DIV><BR>
<DIV>
<P>����������</P><INPUT id=shopnote maxLength=200 
name=shopnote value="<%=rs("shopnote")%>"> ��д�Ļ������ǿ��԰����ƹ��</DIV><BR>
<DIV>
<font 
      color=#000000><b>�Ա�</b></font>
<span class="sex"><span style="WIDTH: 10px">&nbsp;&nbsp;&nbsp;
<select name="sex" size="1" id="select2">
  <option value="1" <%if rs("sex")=1 then%>selected<%end if%>>��</option>
  <option value="0"<%if rs("sex")=0 then%>selected<%end if%>>Ů</option>
</select>
</span></span></DIV><br>
<DIV>
<P>��ϵ��ַ��</P><INPUT id=address maxLength=100 
name=address value="<%=rs("address")%>"></DIV><BR>
<BR>
<BR>
<BR>
<BR>
<BR>
<DIV>
<% Sql2 = "select  count(*) as jiewu  from "&jieducm&"_toushu where username='"&rs("username")&"' "
            Set Rs2 = Server.CreateObject("Adodb.RecordSet")
				Rs2.Open Sql2,conn,1,1
				if NOT rs2.EOF  then
				jiewu=rs2("jiewu")
				end if
				
				  %> 
<P>Ͷ�����˴�����</P><input name="jiewu" type="text" id="jiewu" size="30"  value="<%=jiewu%>" disabled/></DIV>
<BR>
<DIV>
<% Sql2 = "select  count(*) as jiewu1  from "&jieducm&"_toushu where name='"&rs("username")&"' "
            Set Rs2 = Server.CreateObject("Adodb.RecordSet")
				Rs2.Open Sql2,conn,1,1
				if NOT rs2.EOF  then
				jiewu1=rs2("jiewu1")
				end if
				
				  %> 
<P>��Ͷ�ߴ�����</P><input name="jiewu1" type="text" id="jiewu1" size="30"  value="<%=jiewu1%>" disabled/></DIV>
<BR>
<DIV>
<P>���õ�½IP��</P><input name="login_ip" type="text" id="login_ip" size="30"  value="<%=rs("login_ip")%>" />����Ϊ0������½ip</DIV>


<DIV class=regsubmit><INPUT style="CURSOR: pointer; HEIGHT: 30px"  type=submit value="  ȷ��  " > 
&nbsp;  
<INPUT name="��ť"  type=button style="CURSOR: pointer; HEIGHT: 30px" value="  ����  " onClick="history.back();">
</DIV><BR><BR></DIV></DIV>


</FORM></DIV>
</BODY></HTML>
