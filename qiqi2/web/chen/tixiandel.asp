<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->

<!--#INCLUDE FILE="../background.asp"-->

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
id=request.QueryString("id")
Set rs1=server.createobject("ADODB.RECORDSET")
rs1.open "select * from "&jieducm&"_tixian where id="&id&"",conn,3,3
if not (rs1.eof) then
  rs1("start")="0"  
  rs1.update
  rs1.close
  Response.Write("<script>alert('ȷ�����û����ֽ��\n��û�д򵽸��û�֧�����˺���!');window.location.href='admin_bankroll.asp';</script>")
end if
%>