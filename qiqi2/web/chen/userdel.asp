
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
id=request("id")
action=request.form("action")
uid=request.Form("uid")
czm=request.Form("czm")
if action="ok" then
if md5(czm)<>admin_czmpass then

   Set rs=server.createobject("ADODB.RECORDSET")
   rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
   rs.addnew
   rs("username")=session("AdminName")
   rs("class")="ɾ����Ա"
   rs("content")=session("AdminName")&"����Աɾ����Ա�������������"
   rs("jiegou")="ʧ��"
   rs.update
   rs.close
   Response.Write("<script>alert('����������!�벻Ҫ�ظ�������ƽ̨���¼�����Ϊ��');history.back();</script>")
   response.End()
end if

Sql = "select username from "&jieducm&"_register where id="&uid&""
Set Rsm = Server.CreateObject("Adodb.RecordSet")
Rsm.Open Sql,conn,1,1
IF not Rsm.Eof or  not rsm.bof Then
	username=rsm("username")
end if
				
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_register where id in("&id&")",conn,3,3

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_xinyu where username='"&username&"'",conn,3,3

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_zhongxin where username='"&username&"'",conn,3,3

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_jieshou where username='"&username&"'",conn,3,3

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_hei where name='"&username&"'",conn,3,3

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_record where username='"&username&"'",conn,3,3

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_keeper where username='"&username&"'",conn,3,3

   Set rs=server.createobject("ADODB.RECORDSET")
   rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
   rs.addnew
   rs("username")=session("AdminName")
   rs("class")="ɾ����Ա"
   rs("content")=session("AdminName")&"����Աɾ����Ա"&username&"�ɹ�"
   rs("jiegou")="�ɹ�"
   rs.update
   rs.close
   
set rs=nothing
conn.close
set conn=nothing
Response.Write("<script>alert('ɾ����Ա�ɹ���');window.location.href='usermanage.asp';</script>")
response.End()
end if
%><form id="form1" name="form1" method="post" action="">
<input name="action" type="hidden" value="ok" />
<input name="uid" type="hidden" value="<%=id%>" />

  <label>
  �����룺
  <input type="password" name="czm" />
  
  </label>
  <label>
  <input type="submit" name="Submit" value="ȷ��ɾ��" />
  </label>
</form>