<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
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
action=HtmlEncode(trim(request.form("action")))
id=HtmlEncode(trim(request.form("id")))
pushnum=HtmlEncode(trim(request.form("pushnum")))

if action="chongzhi" then
if vip=1 then
pushnum1=int(pushnum)*stupkuanvip
else
pushnum1=int(pushnum)*stupkuan
end if
cunkuan2=cunkuan*100
pushnum2=pushnum1*100

	 if int(cunkuan2)<int(pushnum2) then
			 Response.Write("<script>alert('���Ĵ�����!');window.location.href='mai.asp';</script>")
			 response.End()
	 elseif czm<>request("code") then
		    Response.Write("<script>alert('�����벻��ȷ!');history.back();</script>")
	        response.End()
	elseif int(pushnum)<1 then
			 Response.Write("<script>alert('��������!');history.back();</script>")
	        response.End()
	elseif not isNumeric(pushnum) then
          Response.Write("<script>alert('�����������������룡');history.back();</script>")
          response.End()
     end if	  
	 
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select fabudian,cunkuan From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3  
		rs("fabudian")=rs("fabudian")+pushnum
		rs("cunkuan")=rs("cunkuan")-pushnum1
    	rs.update
    	rs.close  
		 
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_chongjilu " ,Conn,3,3  
	    rs.addnew
		rs("class")="���򷢲���"
		rs("username")=session("useridname")
    	rs("fabudian")=pushnum
    	rs.update
    	rs.close  
		  
	    num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="���򷢲���"
		rs("content")=session("useridname")&"�������߹��򷢲���"&pushnum&"��"
		rs("price")="-"&pushnum1
		rs("jiegou")="����ɹ�"
    	rs.update
    	rs.close    
		call check_jieducm_name(session("useridname"))         
	    Response.Write("<script>alert('��ֵ�ɹ�!');window.location.href='mai.asp';</script>")	
	    response.End()
elseif action="carchong" then
     
	 Sql= "select * from "&jieducm&"_car where id="&id&""
	 Set rs=server.createobject("ADODB.RECORDSET")
	 Rs.Open Sql,conn,3,3
	 if not rs.eof then
	    car_name=rs("car_name")
		car_price=rs("car_price")
	    car_fabudian=rs("car_fabudian")
		rs.close
	else
		 Response.Write("<script>alert('�޴���Ϣ!');history.back();</script>")
        response.End()
	 end if 
	 if int(cunkuan)<int(car_price) then
		Response.Write("<script>alert('���Ĵ�����!');history.back();</script>")
        response.End()
	 end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select jifei,cunkuan,fabudian From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3 
rs("jifei")=rs("jifei")+ car_price
rs("cunkuan")=rs("cunkuan")-car_price
rs("fabudian")=rs("fabudian")+car_fabudian
rs.update
rs.close
	
if tjr<>"" then
 tk=car_price*0.1
 sqlr="update "&jieducm&"_register set cunkuan=cunkuan+"&tk&" where username='"&tjr&"'"
 conn.execute sqlr
num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
rs.addnew
rs("username")=tjr
rs("num")=num
rs("class")="�ƹ㽱��"
rs("content")="���Ƽ��Ļ�Ա"&session("useridname")&"�������߹���"&car_name&",ϵͳ������"&tk&"Ԫ���"
rs("price")=tk
rs("jiegou")="�����ɹ�"
rs.update
rs.close
end if
	
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_chongjilu " ,Conn,3,3  
	    rs.addnew
		rs("class")="����㿨"
		rs("username")=session("useridname")
    	rs("fabudian")=car_fabudian
		rs("cunkuan")=car_price
		rs("jifei")=car_fabudian
    	rs.update
    	rs.close 
		
	   num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="����㿨"
		rs("content")=session("useridname")&"�������߹���"&car_name&",��������"&car_price&"Ԫ,������������"&car_fabudian&"��"
		rs("price")="-"&car_price
		rs("jiegou")="����ɹ�"
    	rs.update
    	rs.close 
call check_jieducm_name(session("useridname"))
call send_message(tjr,"�ƹ㽱��","���Ƽ��Ļ�Ա"&session("useridname")&"������ƽ̨�㿨��ƽ̨������"&tk&"Ԫ���")		 
Response.Write("<script>alert('����ɹ�!ͬʱ��Ҳ�����˻���!');window.location.href='mai.asp';</script>") 
end if	
set rs=nothing
call closeconn()	
%>
			  
