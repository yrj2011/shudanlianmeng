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

dim theid,thdel
theid=request("adid")
thedel=request("del")

	   Set rsr=server.createobject("ADODB.RECORDSET")
	   rsr.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rsr.addnew
		rsr("username")=session("AdminName")
		rsr("class")="ɾ��վ����Ϣ"
		rsr("content")="����Աɾ����վ����Ϣ"
		rsr("jiegou")="�ɹ�"
    	rsr.update
    	rsr.close
if theid="" or thedel="" then
response.write "<script>alert('�������󣬹رմ��ڣ�');window.close();</script>"
response.end
end if

if thedel="user" then
sql="delete from "&jieducm&"_china_user where instr(1,'"&theid&"',uid)>0"
conn.execute(sql) 

sql="delete from "&jieducm&"_china_data where instr(1,'"&theid&"',uid)>0"
conn.execute(sql)

sql="delete from "&jieducm&"_china_message where instr(1,'"&theid&"',uid)>0"
conn.execute(sql)

'''''''''''ɾ��html��ʼ''''''''''''''''

end if

if thedel="data" then
sql="delete from "&jieducm&"_china_data where adid in("&theid&")"
conn.execute(sql) 

sql="delete from "&jieducm&"_china_message where adid in("&theid&")"
conn.execute(sql)

end if

if thedel="message" then
set rs=server.createobject("adodb.recordset")
rs.open "select * from "&jieducm&"_china_message where id in("&theid&") order by adid desc",conn,1,1
dim adid()
a=0
for i=0 to rs.recordcount-1
redim Preserve adid(a)
if ad=rs("adid") then
rs.movenext
else
adid(a)=rs("adid")
ad=rs("adid")
a=a+1
rs.movenext
end if
next
rs.close
set rs=nothing
sql="delete from "&jieducm&"_china_message where id in("&theid&")"
conn.execute(sql)

conn.close
set conn=nothing

end if

if thedel="pass" then
conn.execute("update "&jieducm&"_china_data set mark='yes' where adid in("&theid&")") 
end if
                             
response.redirect request.ServerVariables("HTTP_REFERER")
%>