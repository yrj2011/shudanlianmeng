<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
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
id=request("id")
if request.Form<>"" then
	 code=HtmlEncode(trim(request.form("code")))
	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_register where username='" &session("useridname")& "' and czm='"&code&"'" ,Conn,3,3  
				IF Rs.Eof Then
					Response.Write("<script>alert('�����벻��ȷ!');</script>")
				Else
					 session("czm")=rs("czm")
					 if id="ti" then
					   response.Redirect("remoney.asp")
					 else
    	              response.redirect("PushMission.asp")
					 end if
				end if
set rs=nothing

end if
%>	
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="../images/Default.css" type=text/css rel=stylesheet>
<LINK href="../images/Bbs.css" type=text/css rel=stylesheet>
<LINK href="../images/style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR></HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file=../inc/jieducm_top.asp-->
<SCRIPT language=javascript>
function  save_onclick()
{

    var qq1=document.form.code.value;
  if(qq1=="")
  {
    alert("���������Ĳ����룡");
	document.form.code.focus();
	return false;
	}
	return true;
	}
	</script>
<!-- �����ز˵����� --><!-- �����ز˵�BIG -->
<div align="center">
<DIV 
style="MARGIN-BOTTOM: 10px; WIDTH: 910px; PADDING-TOP: 15px; BACKGROUND-COLOR: #ffffff">
<FORM action="" method="POST" name="form" id="form" onSubmit="return save_onclick()">
<DIV></DIV>

<DIV style="WIDTH: 910px; TEXT-ALIGN: center">
<DIV 
style="PADDING-RIGHT: 20px; PADDING-LEFT: 20px; PADDING-BOTTOM: 20px; WIDTH: 700px; PADDING-TOP: 20px; TEXT-ALIGN: left">
<DIV style="FONT-WEIGHT: bolder; FONT-SIZE: 18px; COLOR: red">���������</DIV><BR>
<DIV style="COLOR: #999900"><FONT 
color=#ff0000>����������������ܼ���������ÿ�ε�½��Ҫ����һ�β��ܷ����������֣����ͷ����㡢�޸ĸ�������</FONT></DIV><BR>
<DIV style="COLOR: red"><SPAN id=ctl00_AllBaseContentPlaceHolder_MsgLabel 
style="COLOR: red"></SPAN></DIV><BR>
<DIV>
<P>�����룺</P><INPUT id=code type=password 
maxLength=16 name=code><SPAN 
id=ctl00_AllBaseContentPlaceHolder_RequiredFieldValidator1 
style="DISPLAY: none; COLOR: red">* �����벻��Ϊ�գ�</SPAN></DIV><BR>
<DIV><INPUT id=ctl00_AllBaseContentPlaceHolder_button1  type=submit value="  �ύ  " name=ctl00$AllBaseContentPlaceHolder$button1></DIV><BR><BR></DIV></DIV>


<DIV> </DIV>



</FORM></DIV></div>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
call closeconn()%>
</BODY></HTML>
