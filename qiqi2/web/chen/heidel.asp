<!--#include file="inc/conn.asp"-->

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
id=request("id")
if id="" then
 Response.Write("<script>alert('����ѡ��Ҫɾ���ļ�¼!');history.back();</script>")
 response.End()
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_hei where id in ("&id&")",conn,3,3
set rs=nothing
		Set rsr=server.createobject("ADODB.RECORDSET")
	  rsr.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rsr.addnew
		rsr("username")=session("AdminName")
		rsr("class")="ɾ��������"
		rsr("content")="����Աɾ���˺�����"		
		rsr("jiegou")="�ɹ�"
    	rsr.update
    	rsr.close
conn.close
set conn=nothing
  response.Redirect("hei.asp")
  %>