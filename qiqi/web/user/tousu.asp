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
'Copyright (C) 2008��2011 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
action=HtmlEncode(trim(request.form("action")))
pid=request("pid")
usernamef=request("usernamet")

if action="ok" then
title=HtmlEncode(trim(request.form("title")))
name1=HtmlEncode(trim(request.form("name")))
pid=HtmlEncode(trim(request.form("pid")))
content=(trim(request.form("content")))
classid=request.form("class")
if classid="0" then
	   Response.Write("<script>alert('����ѡ������ԭ��!');history.back();</script>")
	   response.end()
elseif content="" then
	   Response.Write("<script>alert('����ԭ��֤�ݲ���Ϊ��!');history.back();</script>")
	   response.end()
end if
   Sql = "select * from "&jieducm&"_register where username='"&name1&"'"
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	IF Rs.Eof Then
	   Response.Write("<script>alert('�Բ����޴��û�!');history.back();</script>")
	   response.end()
	end if

        Set rs=server.createobject("ADODB.RECORDSET")
        rs.open "Select * From "&jieducm&"_toushu where pid='"&pid&"'" ,Conn,3,3  
		if not rs.eof then
		  Response.Write("<script>alert('���������ύ����!');history.back();</script>")
	      response.end() 
		else
	    rs.addnew
		rs("title")=title
		rs("name")=name1
    	rs("pid")=pid
		rs("content")=content
		rs("username")=session("useridname")
		rs("pid")=pid
		rs("bianjie")="��"
		rs("bianjie2")="��"
		rs("guang")="��"
		rs("class")=classid
    	rs.update
    	rs.close
		end if
		
		
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="����"
		rs("content")=session("useridname")&"����"&name1&"�ɹ�"
		rs("jiegou")="�ɹ�"
    	rs.update
    	rs.close	
		set rs=nothing
		Response.Write("<script>alert('���߳ɹ�!');window.location.href='complaintmy.asp';</script>")
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
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</HEAD>
<SCRIPT language=javascript>
function  save_onclick12()
{

    var Price=document.formf.title.value;
  if(Price=="")
  {
    alert("��������⣡");
	document.formf.title.focus();
	return false;
	}

    var qu=document.formf.qu.value;
  if(qu=="")
  {
    alert("��������������");
	document.formf.qu.focus();
	return false;
	}

 var ProUrl=document.formf.name.value;
  if(ProUrl=="")
  {
    alert("�����뱻�����ˣ�");
	document.formf.name.focus();
	return false;
	}
  
  var Shopkeeper=document.formf.pid.value;
  if(Shopkeeper=="")
  {
    alert("����������ID��");
	document.formf.pid.focus();
	return false;
	}
	if(!Number(Shopkeeper))
	  
  {
    alert("����ID����������!!");
	document.formf.pid.focus();
	return false;
	}
 var content=document.formf.content.value;
  if(content=="")
  {
    alert("���ݲ���Ϊ�գ�");
	
	return false;
	}
	
  return true;  
}
</SCRIPT>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   
<!--#include file=../inc/jieducm_top.asp-->
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
                    <div class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt;��Ҫ���� &gt;&gt; </div>
                    <div class=pp8><strong>��Ҫ����</strong></div>
					<FORM name=formf method=post action=""  onSubmit="return save_onclick12()">
					<input name="action" type="hidden" value="ok">
					<table width="680" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="110" height="40" align="center">��&nbsp;&nbsp;&nbsp;�⣺</td>
              <td width="490"><input name="title" type="text" id="title" size="25" /> 
              (ֻ��������е���������ɵ���������) </td>
            </tr>
            <tr>
              <td height="40" align="center">����ԭ��</td>
              <td><select name="class">			   
               <option value="0">����ѡ������ԭ��</option>
                <option value="���δ���Ա�֧����ȴ��ƽ̨�����Ѿ�֧����">���δ���Ա�֧����ȴ��ƽ̨�����Ѿ�֧����</option>
                <option value="���δ���Ա�ȷ���ջ�������ȴ��ƽ̨�������Ѻ���">���δ���Ա�ȷ���ջ�������ȴ��ƽ̨�������Ѻ���</option>
                <option value="���Ҳ���2Сʱ�󣬶Է���δ����һ������(��������)">���Ҳ���2Сʱ�󣬶Է���δ����һ������(��������)</option>
				<option value="�����ջ�ʱ��12Сʱ�������δ�ջ�����(ʵ������)">�����ջ�ʱ��12Сʱ�������δ�ջ�����(ʵ������)</option>
				
				<option value="������Ա�����12Сʱ��������δ����(ʵ������)">������Ա�����12Сʱ��������δ����(ʵ������)</option>			
				<option value="�����ƽ̨ȷ�Ϻ���12Сʱ��������δȷ�Ϻ���(ʵ������)">�����ƽ̨ȷ�Ϻ���12Сʱ��������δȷ�Ϻ���(ʵ������)</option>
				<option value="δ��ȷ���ջ�ʱ�䣬���ȴ��ǰȷ���ջ�">δ��ȷ���ջ�ʱ�䣬���ȴ��ǰȷ���ջ�</option>
				<option value="��Ҹ��˽�����Ϣ�����ƻ��ջ�����Ϣ��д���淶">��Ҹ��˽�����Ϣ�����ƻ��ջ�����Ϣ��д���淶</option>
				<option value="��Ĭ��������Ʒ�����ȴ�������������">��Ĭ��������Ʒ�����ȴ�������������</option>
				<option value="��ҵ�ƽ̨���ֺ����Ա���Ų�һ��">��ҵ�ƽ̨���ֺ����Ա���Ų�һ��</option>
				
				<option value="���δ�����ҵ�Ҫ��д����">���δ�����ҵ�Ҫ��д����</option>
				<option value="����û�����Ž�����">����û�����Ž�����</option>
				<option value="��������˿�">��������˿�</option>
				<option value="���ҷ����ײ�����">���ҷ����ײ�����</option>
				<option value="�������">�������</option>
				
              </select></td>
            </tr>
            <tr>
              <td height="40" align="center">�������ˣ�</td>
              <td><input name="name" type="text" id="name" size="25"  value="<%=usernamef%>"/></td>
            </tr>
            <tr>
              <td height="40" align="center">��&nbsp;��&nbsp;ID��</td>
              <td><input name="pid" type="text" id="pid" size="25"  value="<%=pid%>"/></td>
            </tr>
            <tr>
              <td height="40" align="center">����ԭ��֤�ݣ�</td>
              <td>
               <textarea name="Content" id="content" style="display:none"></textarea> 
            
              <iframe id="Editor" name="Editor" src="../HtmlEditor/index.html?ID=content" frameborder="0" marginheight="0"     marginwidth="0" scrolling="No" style="height:320px;width:100%"></iframe>  <span class="STYLE1">
              <p>���ϴ���������Ա���ͼ���Ȱѽ�ͼ�ϴ����Ա�ͼƬ�ռ䣬Ȼ���Ƶ�ַ���ɣ�</p></span>
              </td>
            </tr>
            <tr>
              <td height="70">&nbsp;</td>
              <td valign="middle">&nbsp;&nbsp;&nbsp;
                <input type="submit" name="button" id="button" value="�ύ" style="FONT-WEIGHT: bold; WIDTH: 120px; CURSOR: pointer; COLOR: #000000; HEIGHT: 26px">
                <input type="reset" name="button2" id="button2" value="����" style="FONT-WEIGHT: bold; WIDTH: 120px; CURSOR: pointer; COLOR: #000000; HEIGHT: 26px"></td>
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
rs.close
set rs=nothing
call closeconn()%>
</BODY></HTML>
