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
      '发送短信成功,可进一步做成功后续处理
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
 Response.Write("<script>alert('请不要重复绑定!平台会记录你的行为！');history.back();</script>")
 response.End()
end if

       Set rs=server.createobject("ADODB.RECORDSET")
	   rs.open "Select * From "&jieducm&"_register where dst="&dst&"" ,Conn,3,3 
	   if (not rs.bof) and  (not rs.eof) then
	         Response.Write("<script>alert('对不起此手机号已被其它用户绑定!');history.back();</script>")
	        response.End()		
	   end if  

	 if czm<>request("code") then
		    Response.Write("<script>alert('操作码不正确!');history.back();</script>")
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
		rs("class")="进行手机认证"
		rs("content")=session("useridname")&"进行手机认证接收手机号："&dst&"时间："&now()&"ip:"&Request.ServerVariables("REMOTE_ADDR")&""
		rs("price")=""
		rs("jiegou")="等待手机接收验证码"
    	rs.update
    	rs.close     
'///////////////////////////////////////
'文件名：sendsms3.asp
'功能：示范调用接口发送短信
'编写：魏东（wei_dong）
'最后更新日期：2005-10-12
'//////////////////////////////////////

'-----------------------------------------------------------------
'名称：SendSms(UserName, UserPass, DstMobile, SmsMsg)
'参数：
'    UserName: 短信王用户名
'    UserPass： 短信王密码
'    DstMobile:目标手机号码
'    SmsMsg:发送内容
'功能：发送短信
'返回值：
'       true,已经成功发送
'       false,未成功发送  

'-----------------------------------------------------------------

'////测试程序主体

sDst=dst  '接收短信的目标手机号码
smsg=HtmlEncode(trim(request.form("msg")))'短信内容
smsg=smsg&ranNum

if SendSms(smsname, smspwdphone, sDst, smsg) then
   '///发送成功后的处理，例如扣费，记录用户发送短信记录等
    Response.Write("<script>alert('发送成功!');window.location.href='Center_Userlock.asp?action=ok&dst="&sdst&"';</script>")	  
else
   response.write "发送失败!"
end if

else
	Response.Write("<script>alert('请输入手机号!');history.back();</script>")	
	response.End()
end if
%>
