<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#INCLUDE FILE="inc/md5.asp"-->

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
Set rs=server.createobject("ADODB.RECORDSET")
if request.Form<>"" then
  if request("password")<>"" then
	 password=md5(Request.Form("password"))
  end if
	 mail=request.Form("email")
     HomePage=Trim(Request.Form("HomePage"))
	 qq=request.Form("qq")
	 face=request.Form("face")
	 sex=request.Form("sex")
	 tjr=Request.Form("tjr")
	 phone=Request.Form("phone")
	 rname=Request.Form("rname")
	 shopnote=Request.Form("shopnote")
	 address=Request.Form("address")
	 czm=Request.Form("czm")
	 weiti=request("weiti")
	 daai=request("daai")
	 dst=request("dst")
	 login_ip=request.Form("login_ip")
	 if homepage="http://" then homepage="" 
	  rs.open "Select * From "&jieducm&"_register where username='"&request("username")&"' " ,Conn,3,3  
	  if request("password")<>"" then
		rs("password1")=password
	  end if
    	rs("mail")=mail
		rs("homepage")=homepage
		rs("qq")=qq
		rs("sex")=sex
		rs("tjr")=tjr
		rs("phone")=phone
		rs("rname")=rname
		rs("shopnote")=shopnote		
		rs("address")=address
		rs("czm")=czm
		rs("weiti")=weiti
		rs("daai")=daai
		rs("login_ip")=login_ip
		rs("dst")=dst
    	rs.update
    	rs.close
		
		
	Set rsr=server.createobject("ADODB.RECORDSET")
	  rsr.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rsr.addnew
		rsr("username")=session("AdminName")
		rsr("class")="�޸��û���Ϣ"
		rsr("content")=session("AdminName")&"�޸���"&request("username")&"����Ϣ"		
		rsr("jiegou")="�޸ĳɹ�"
    	rsr.update
    	rsr.close
		username=request("username")
     Response.Write("<script>alert('��Ϣ�޸ĳɹ�!');window.location.href='usermanage.asp?action=sear&text="&username&"';</script>")

set rs=nothing
conn.close
set conn=nothing
end if
%>