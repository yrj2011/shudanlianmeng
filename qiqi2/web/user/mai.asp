<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->
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

    var Price=document.form.PushNum.value;
  if(Price=="")
  {
    alert("�����빺�������");
	document.form.PushNum.focus();
	return false;
	}
  if(Price<1)
	  
  {
    alert("���������1����!");
	document.form.PushNum.focus();
	return false;
	}
  if(!Number(Price))	  
  {
    alert("������������!");
	document.form.PushNum.focus();
	return false;
	}
if   ((document.form.PushNum.value.indexOf("-")   ==   0)||!(document.form.PushNum.value.indexOf(".")   ==   -1)){   
  alert("����ΪС������");   
  document.form.PushNum.focus();   
  return   false;   
  }
	 var code=document.form.code.value;
  if(code=="")
  {
    alert("�����벻��Ϊ�գ�");
	document.form.code.focus();
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
                    <div class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt; ���򷢲��� &gt;&gt; </div>
                    <div class=pp8><strong>���򷢲���</strong></div>
                    <br>
                   
					 <DIV style="CLEAR: both; WIDTH: 710px;  BACKGROUND-COLOR: #f3f8fe">
<DIV 
style="CLEAR: both; PADDING-RIGHT: 30px; PADDING-LEFT: 30px; HEIGHT: 100%">

<DIV> </DIV>
<DIV style="CLEAR: both; PADDING-LEFT: 30px; BACKGROUND: url(/images/shangyi.gif) no-repeat left 50%; PADDING-BOTTOM: 5px; COLOR: #004000; PADDING-TOP: 7px; BORDER-BOTTOM: #abbec8 1px dashed"><SPAN 
style="COLOR: red">��ֵ���񿨹���(ȫ�Զ�������) ע�����÷����㣬��Ҳ����ȥ�����񣬲�һ����Ҫ����ģ�</SPAN><SPAN 
id=ctl00_AllBaseContentPlaceHolder_MsgLabel 
style="FONT-SIZE: 20px; COLOR: red"></SPAN></DIV>
<DIV style="CLEAR: both; PADDING-LEFT: 30px; PADDING-BOTTOM: 5px; COLOR: #004000; LINE-HEIGHT: 70px; PADDING-TOP: 7px; BORDER-BOTTOM: #abbec8 1px dashed; ">
  <div align="center"><a href="md5_pay.asp" >
  <IMG onMouseOver="this.src ='../images/pay_link.gif'" onMouseOut="this.src='../images/pay.gif'" src="../images/pay.gif" border=0></A></div>
</DIV>
<DIV style="CLEAR: both; PADDING-LEFT: 30px; PADDING-BOTTOM: 5px; PADDING-TOP: 15px">
<DIV style="FLOAT: left; WIDTH: 350px"><IMG 
src="../images/jieducm_258shua.gif"></DIV>
  <form name="form" method="post" action="chongzhi.asp" id="form"  onSubmit="return save_onclick()">
  <input name="action" type="hidden" value="chongzhi">
<DIV style="FLOAT: left; WIDTH: 250px; LINE-HEIGHT: 150%; PADDING-TOP: 20px">
<SPAN style="COLOR: red">ӵ�з����㣬�Ϳ��Է����Լ������񣬻�ú����� ��Ҫ���������������������������򷢲��㣡<br>vip��Ա
<%if stupkuanvip <1 then
response.Write("0")
end if
response.Write(stupkuanvip)
%>
Ԫ/��,��ͨ��Ա:
<%if stupkuan <1 then
response.Write("0")
end if
response.Write(stupkuan)
%>Ԫ/��</SPAN><BR>

����������<INPUT id=PushNum style="WIDTH: 30px" value=10 name=PushNum onKeyUp="if(isNaN(value))execCommand('undo')">
��/ÿ��
<%
if vip=1 then
if stupkuanvip<1 then
response.Write("0")
end if
response.Write(stupkuanvip)
else
if stupkuan<1 then
response.Write("0")
end if
response.Write(stupkuan)
end if
%>Ԫ<BR>
�����룺<INPUT name=code type="password" id=code size="10">  <A style="COLOR: red; TEXT-DECORATION: underline" class=renwu-link  href="../user/paypoint.asp">ȥ�����г�����</A>
<BR>

<INPUT id=ImageButton1  onClick="return confirm('��ȷ�����򷢲�����');" type=image src="../images/buy1.png" name=ImageButton1></DIV>
</form>
</DIV>
<%
   	sql="SELECT * FROM "&jieducm&"_car  order by car_sort"
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	if not rs.eof then
	do while not rs.eof
%>
<DIV style="CLEAR: both; PADDING-LEFT: 30px; PADDING-BOTTOM: 5px; PADDING-TOP: 15px">
<DIV style="FLOAT: left; WIDTH: 350px"><IMG src="../<%=rs("car_pic")%>"></DIV>
<form name="Form2" method="post" action="chongzhi.asp" id="Form2">
<input name="action" type="hidden" value="carchong">
<input name="id" type="hidden" value="<%=rs("id")%>">
<DIV style="FLOAT: left; WIDTH: 200px; LINE-HEIGHT: 150%; PADDING-TOP: 20px">
<SPAN style="COLOR: red"><%=rs("car_content")%></SPAN><BR>
�����㣺<%=rs("car_fabudian")%>��<BR>
<%=rs("car_name")%>��<%=rs("car_price")%>Ԫ<BR>
<INPUT id=ImageButton2  onClick="return confirm('��ȷ��Ҫ����<%=rs("car_name")%>��');" type=image src="../images/buy1.png" name=ImageButton2>
</DIV>
</form>
</DIV>
<%
Rs.MoveNext
Loop
end if%>

<DIV></DIV>
</DIV></DIV>
                  </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%call closeconn()%>
</BODY></HTML>
