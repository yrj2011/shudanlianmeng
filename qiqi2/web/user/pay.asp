<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
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
'Copyright (C) 2008��2009 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
From_url = Cstr(Request.ServerVariables("HTTP_REFERER"))
Serv_url = Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(From_url,8,len(Serv_url)) <> Serv_url or Request.ServerVariables("HTTP_REFERER")="" then
response.write "����ʱ,��رձ�ҳ�沢���µ�½����!"
response.end
end If
action=request.form("action")
if action="1" then
num=Replace(Trim(Request.Form("num")),"'","''")
price=Replace(Trim(Request.Form("price")),"'","''")
username=Replace(Trim(Request.Form("username")),"'","''")

if int(fabudian)<int(stuppaynum) then
 	Response.Write("<script>alert('���������"&stuppaynum&"���ܳ���!');history.back();</script>")
	response.End()	
elseif(int(fabudian)<int(num)) then
    Response.Write("<script>alert('����۵ķ����㲻�ܴ��������еķ�����!');history.back();</script>")
	response.End()	
elseif czm<>request.form("czm") then
	Response.Write("<script>alert('�����벻��ȷ!');history.back();;</script>")
	response.End()	
elseif stuppayis=0 then
    Response.Write("<script>alert('�˹����ѹرգ�');history.back();</script>")
   response.End()
elseif jifei< stuppayjifei	then
    Response.Write("<script>alert('���ָ���"&stuppayjifei&"�ſ���ʹ�ô˹��ܣ�');history.back();</script>")
   response.End()
elseif int(num)<1 then
	Response.Write("<script>alert('������������������!');history.back();;</script>")
	response.End() 
elseif not isNumeric(num) then
   Response.Write("<script>alert('�����������������룡');history.back();</script>")
   response.End()
elseif not isNumeric(price) then
   Response.Write("<script>alert('�����������������룡');history.back();</script>")
   response.End()
elseif int(price)<1 then
   Response.Write("<script>alert('�����������������룡');history.back();</script>")
   response.End()
end if

if username<>"" then
     Set rs=server.createobject("ADODB.RECORDSET")
	 rs.open "Select * From "&jieducm&"_register where username='"&username&"'" ,Conn,1,1  
	   if rs.eof then
	       Response.Write("<script>alert('ָ������û���������');history.back();</script>")
		   response.End()
       elseif session("useridname")=username then
	      Response.Write("<script>alert('ָ������û����������Լ�!');history.back();</script>")
		  response.End()
	  end if
end if	  

if vip=1 then
jieducm_shou=(stupvippaynum/100)*num
else
jieducm_shou=(stuppupaynum/100)*num
end if

if (num+jieducm_shou)>fabudian then
    Response.Write("<script>alert('����۵ķ����㲻�ܴ��������еķ�����!');history.back();</script>")
	response.End()	
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select fabudian From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
if rs.eof then
 Response.Write("<script>alert('�޴˻�Ա');window.location.href='pay.asp';</script>")
response.End()
else
rs("fabudian")=rs("fabudian")-num-jieducm_shou
rs.update
rs.close
end if	  

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_pay " ,Conn,3,3  
rs.addnew
rs("jieducm_username")=session("useridname")
rs("jieducm_num")=num
rs("jieducm_price")=price
rs("jieducm_maijia")=username
rs("jieducm_datatime")=now()
rs("jieducm_start")=0
rs.update
rs.close
	    nums=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=nums
		rs("class")="���۷�����"
		rs("content")=session("useridname")&"�������߳��۷�����"&num&"��,���ó����ܼ۸�Ϊ"&price&"Ԫ��������"&jieducm_shou&"��������"
		rs("price")=0
		rs("jiegou")="���۳ɹ�"
    	rs.update
    	rs.close  
		
 Response.Write("<script>alert('�����ɹ���');window.location.href='../user/paypoint.asp';</script>")
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
    var num=document.form.num.value;
  if(num=="")
  {
    alert("������Ҫ���۵ķ�����������");
	document.form.num.focus();
	return false;
	}
 if(!Number(num))
	  
  {
    alert("������������!");
	document.form.num.focus();
	return false;
	}
if ((document.form.num.value.indexOf("-")   ==   0)||!(document.form.num.value.indexOf(".")   ==   -1)){   
  alert("������������ΪС������");   
  document.form.num.focus();   
  return   false;   
  }
 
  var price=document.form.price.value;
  if(price=="")
  {
    alert("��������۵��ܼ۸�");
	document.form.price.focus();
	return false;
	}
	  
 var czm=document.form.czm.value;
  if(czm=="")
  {
    alert("���������Ĳ����룡");
	document.form.czm.focus();
	return false;
	}
  
  return true;  
}
</script>   
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
                    <div class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt;���۷����� &gt;&gt; </div>
                    <div class=pp8><strong>���۷�����</strong></div>
                    <div style="MARGIN-TOP: 10px; LINE-HEIGHT: 200%">
                      <ul>
                        <li>
<SPAN style="COLOR: red">*����ȡ��������Ϊ"̯��",ȡ�����۲����� ������,�붨λ�ú��ʵļ۸��ٳ���!������������,��������ͨ���ٷ��һ�ֱ�ӻ�ȡ�ֽ�!</SPAN></li>
                      </ul>
                    </div><br> 
                    <form action="" method="POST" name="form" id="form" onSubmit="return save_onclick()">
                      <table width="679" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="35" align="right"><div align="right"></div></td>
                          <td> ����ǰ���÷������ǣ�<%=fabudian%> �� </td>
                        </tr>
                        <tr>
                          <td width="123" height="35" align="right"><div align="right">Ҫ���۵ķ�����������</div></td>
                          <td width="556"><input  name=num id=num  onKeyUp="if(isNaN(value))execCommand('undo')" maxlength="4">
                            �� ӵ��<%=stuppaynum%>����������ܳ���!&nbsp; vip��Ա��<%=stupvippaynum%>%�����ѡ���ͨ��Ա��<%=stuppupaynum%>%������</td>
                        </tr>
                        <tr>
                          <td height="35"><div align="right">���ó��۵��ܼ۸�</div></td>
                          <td><input type="text" name="price" id="price"  onKeyUp="if(isNaN(value))execCommand('undo')">
                          ���ۼ۸�Խ�͹������Խ��</td>
                        </tr>
                        <tr>
                          <td height="35"><div align="right">ָ������û�����</div></td>
                          <td><label>
                            <input name="username" type="text" id="username">
                          ��ɷ�ָ����ĳһ���û��ſ��Թ�����Щ������</label></td>
                        </tr>
                        <tr>
                          <td height="35"><div align="right">�����룺</div></td>
                          <td><label>
                            <input name="czm" type="password" id="czm">
                          </label></td>
                        </tr>
                        <tr>
                          <td height="35">&nbsp;</td>
                          <td><input type="submit" name="button" id="button" value="ȷ�ϳ��۷�����" <%if stuppayis=0 or jifei<stuppayjifei then%> disabled <%end if%> />
                              <input type="hidden" name="action" value="1">
                              &nbsp;&nbsp;  &nbsp; ���ָ���<%=stuppayjifei%>��ſ���ʹ��</td>
                        </tr>
                      </table>
                    </form>
            
					
				  </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close()
set rs=nothing
call closeconn()%>
</BODY></HTML>
