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
id=request.QueryString("id")
zfbnum=request("zfbnum")
action=request("action")
if action="ok" then
Set rs1=server.createobject("ADODB.RECORDSET")
rs1.open "select * from "&jieducm&"_tixian where id="&id&"",conn,3,3
if not (rs1.eof) then
price=rs1("ReNum")
username=rs1("username")

  rs1("start")="1"  
  rs1("zfbnum")=zfbnum
  rs1.update
  rs1.close
  
  	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rsr=server.createobject("ADODB.RECORDSET")
	  rsr.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rsr.addnew
		rsr("username")=username
    	rsr("num")=num
		rsr("class")="��������"
		rsr("content")="����Աȷ������ "&price&"Ԫ"
		rsr("price")=price
		rsr("jiegou")="���ֳɹ�"
    	rsr.update
    	rsr.close
  Response.Write("<script>alert('ȷ�����û����ֽ��\n�Ѵ򵽸��û�֧�����˺���!');window.location.href='admin_bankroll.asp';</script>")
end if

elseif action="rem" then

       Set Rs = Server.CreateObject("Adodb.RecordSet")
	   rs.open "Select * From "&jieducm&"_tixian where  id="&id&" and start<>'2'"  ,Conn,3,3  
	   if not rs.eof then
        rs("start")=2
		username=rs("username")
		price=rs("ReNum")
    	rs.update
    	rs.close
		
 			  	Sql3 = "select * from "&jieducm&"_register where username='"&username&"'"
				Set Rs3 = Server.CreateObject("Adodb.RecordSet")
				Rs3.Open Sql3,conn,3,3
				if not(rs3.eof) then
				  rs3("cunkuan")=rs3("cunkuan")+price
				  rs3.update
				  rs3.close
				end if
				
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rsr=server.createobject("ADODB.RECORDSET")
	  rsr.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rsr.addnew
		rsr("username")=username
    	rsr("num")=num
		rsr("class")="��������"
		rsr("content")="����Ա���г������� "&price&"Ԫ"
		rsr("price")="+"&price
		rsr("jiegou")="�����ɹ�"
    	rsr.update
    	rsr.close
  Response.Write("<script>alert('���ѳɹ��������û�����,\��Ӧ����ѷ��������û��Ĵ����!');window.location.href='admin_bankroll.asp';</script>")
  end if
elseif action="shou" then
       Set Rs = Server.CreateObject("Adodb.RecordSet")
	   rs.open "Select * From "&jieducm&"_tixian where  id="&id&" and start<>'2'"  ,Conn,3,3  
	   if not rs.eof then
        rs("start")=3
    	rs.update
    	rs.close
		 Response.Write("<script>alert('���ѳɹ��������û�����������!');window.location.href='admin_bankroll.asp';</script>")
       end if
elseif action="del" then

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_tixian where id in ("&id&")",conn,3,3
set rs=nothing
conn.close
set conn=nothing
Response.Write("<script>alert('��ɾ���˴�����!');window.location.href='admin_bankroll.asp';</script>")

end if
%>