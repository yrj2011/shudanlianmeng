 <!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="../md5.asp"-->
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
jieducm_ka=request("jieducm_ka")
if session("AdminName")="" then
   response.redirect("admin_login.asp")
end if
if request.form<>"" then
key=request("key")
czm=request("czm")
		
if md5(czm)<>admin_czmpass then
     Set rsr=server.createobject("ADODB.RECORDSET")
	  rsr.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rsr.addnew
		rsr("username")=session("AdminName")  
		rsr("class")="��ֵ�������ѯ"
		rsr("content")=session("AdminName")&"��ѯ����Ϊ"&key&"������"		
		rsr("jiegou")="ʧ��"
    	rsr.update
    	rsr.close
 Response.Write("<script>alert('����������!�벻Ҫ�ظ�����!ƽ̨���¼�����Ϊ!');window.history.back();</script>")
  response.end()
end if

Set rsr=server.createobject("ADODB.RECORDSET")
rsr.open "Select * From "&jieducm&"_beifei where ka='"&key&"'" ,Conn,1,1
if not rsr.eof then
kahao=rsr("ka")
pwd=rsr("pwd")
else
 Response.Write("<script>alert('�޴˿���!');window.history.back();</script>")
  response.end()
end if

     Set rsr=server.createobject("ADODB.RECORDSET")
	  rsr.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rsr.addnew
		rsr("username")=session("AdminName")  
		rsr("class")="��ֵ�������ѯ"
		rsr("content")=session("AdminName")&"��ѯ����Ϊ"&key&"������"		
		rsr("jiegou")="�ɹ�"
    	rsr.update
    	rsr.close
end if
%>
<html>
<head>
<title>��ֵ������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Inc/Admin.css" type="text/css">

</head>
<body>
<table align="center" width="100%" border="1" cellspacing="0" cellpadding="4" class=dj586_Com_bk style="border-collapse: collapse">
<tr class=dj586_Com_ss> 
<td colspan="6"><div class='bodytitle'><div class='bodytitleleft'></div><div class='bodytitletxt'>��ֵ������ </div></div></td>
</tr>
<tr align="left" class=dj586_Com_ds>
<td colspan="6">  ��������<a href=Admin_Okjicar.asp>�㿨��ֵ��¼</a>| <a href="Admin_Jicar.asp">��ֵ������</a> | </td>      
</tr>
</table><br><form name="myform" method="post" action="">
<table width="619" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="73">�����뿨�ţ�</td>
    <td width="164">
      <input type="text" name="key" id="key" value="<%=jieducm_ka%>">    </td>
    <td width="61">�����룺</td>
    <td width="164"><input type="password" name="czm" id="czm"></td>
    <td width="157"><input type="submit" name="button" id="button" value="��ѯ����" onClick="return CheckForm();"></td>
  </tr>
</table> 
</form>
<table width="400" border="0">
  <tr>
    <td>����</td>
    <td>����</td>
  </tr>
  
  <tr>
    <td><%=kahao%></td>
    <td><%=pwd%></td>
  </tr>
</table>

</body>
</html>
