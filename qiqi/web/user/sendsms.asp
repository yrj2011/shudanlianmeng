<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->
<%
Function SendSms(UserName, UserPass, DstMobile, SmsMsg) 

	set http = Server.CreateObject("Microsoft.XMLHTTP")
	http.Open "GET", "http://www.china-sms.com/send/gsend.asp?name="&UserName&"&pwd="&UserPass&"&dst="&DstMobile&"&msg="&SmsMsg&"", false 
	http.setRequestHeader "Content-type:", "text/xml;charset=GB2312"
	'http.setRequestHeader "Charset", "GB2312"
	http.Send
	msg=http.ResponseText
	set	http = nothing
	
	if left(msg,4)<>"id=0" then 
      '���Ͷ��ųɹ�,�ɽ�һ�����ɹ���������
      SendSms=true
      exit Function
    else
      response.write msg
      SendSms=false
	  exit Function
    end if
End Function
action=HtmlEncode(trim(request.form("action")))
if action="ok" then
dst=HtmlEncode(trim(request.form("dst")))

if dst1<>0 then
 Response.Write("<script>alert('�벻Ҫ�ظ���!ƽ̨���¼�����Ϊ��');history.back();</script>")
 response.End()
end if

       Set rs=server.createobject("ADODB.RECORDSET")
	   rs.open "Select * From "&jieducm&"_register where dst="&dst&"" ,Conn,3,3 
	   if (not rs.bof) and  (not rs.eof) then
	         Response.Write("<script>alert('�Բ�����ֻ����ѱ������û���!');history.back();</script>")
	        response.End()		
	   end if  

	 if czm<>request("code") then
		    Response.Write("<script>alert('�����벻��ȷ!');history.back();</script>")
	        response.End()
	 else
	   randomize
	   ranNum=int(900000*rnd)+100000
	   Set rs=server.createobject("ADODB.RECORDSET")
	   rs.open "Select * From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3     
		rs("codenum")=ranNum
		rs("dst1")=request.Form("dst")
    	rs.update
    	rs.close	
     end if
	 
	    num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="�����ֻ���֤"
		rs("content")=session("useridname")&"�����ֻ���֤�����ֻ��ţ�"&dst&"ʱ�䣺"&now()&"ip:"&Request.ServerVariables("REMOTE_ADDR")&""
		rs("price")=""
		rs("jiegou")="�ȴ��ֻ�������֤��"
    	rs.update
    	rs.close     
'///////////////////////////////////////
'�ļ�����sendsms3.asp
'���ܣ�ʾ�����ýӿڷ��Ͷ���
'��д��κ����wei_dong��
'���������ڣ�2005-10-12
'//////////////////////////////////////

'-----------------------------------------------------------------
'���ƣ�SendSms(UserName, UserPass, DstMobile, SmsMsg)
'������
'    UserName: �������û���
'    UserPass�� ����������
'    DstMobile:Ŀ���ֻ�����
'    SmsMsg:��������
'���ܣ����Ͷ���
'����ֵ��
'       true,�Ѿ��ɹ�����
'       false,δ�ɹ�����  

'-----------------------------------------------------------------

'////���Գ�������

sDst=dst  '���ն��ŵ�Ŀ���ֻ�����
smsg=HtmlEncode(trim(request.form("msg")))'��������
smsg=smsg&ranNum

if SendSms(smsname, smspwdphone, sDst, smsg) then
   '///���ͳɹ���Ĵ�������۷ѣ���¼�û����Ͷ��ż�¼��
    Response.Write("<script>alert('���ͳɹ�!');window.location.href='Center_Userlock.asp?action=ok&dst="&sdst&"';</script>")	  
else
   response.write "����ʧ��!"
end if

else
	Response.Write("<script>alert('�������ֻ���!');history.back();</script>")	
	response.End()
end if
%>
