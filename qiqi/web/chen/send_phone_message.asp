<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����վ����Ϣ</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>
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



if request.Form<>"" then
messagename=request("messagename")
messagelxfs=request("messagelxfs")
messagetext=request("messagetext")
uid=HtmlEncode(trim(request.form("ggwei")))

set Rsstup=server.createobject("adodb.recordset")
sql="select MailServer,MailServerUserName,MailServerPassWord,MailDomain from "&jieducm&"_stup "
Rsstup.open sql,conn,1,1
if not(Rsstup.bof) then
 stupMailServer=rsstup("MailServer")
 stupMailServerUserName=rsstup("MailServerUserName")
 stupMailServerPassWord=rsstup("MailServerPassWord")
 stupMailDomain=rsstup("MailDomain")
end if

if uid<>"all" then
 	    Set rs=server.createobject("ADODB.RECORDSET")
        rs.open "Select mail  From "&jieducm&"_register where username='"&uid&"'" ,Conn,3,3 
		if rs.eof then
		   response.write("<script>alert('�Բ���,�޴��û�');history.back();</script>")
		   response.End()
		else
		 if SendSms(smsname, smspwdphone, jdst, message) then


Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select cunkuan From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
if rs.eof then
 Response.Write("<script>alert('���ܳ�ʱ�����');history.back();</script>")
 response.End()
else
rs("cunkuan")=rs("cunkuan")-price
rs.update
rs.close
end if

		end if
else

        Set rs=server.createobject("ADODB.RECORDSET")
        rs.open "Select mail  From "&jieducm&"_register " ,Conn,3,3 
		if rs.eof then
		   response.write("<script>alert('�Բ���,��������');history.back();</script>")
		   response.End()
		else
		do while not rs.eof
		  semail=rs("mail")
		  title=messagelxfs '�ʼ�����
          book=messagetext '�ʼ�����
          Set JMail = Server.CreateObject("JMail.Message") 
          JMail.Charset = "gb2312"
          JMail.From = stupMailDomain 'SMTP����
          JMail.FromName = messagename '������
          JMail.Subject = title 
          JMail.MailServerUserName = stupMailServerUserName'�û���
          JMail.MailServerPassword = stupMailServerPassWord'����
          JMail.Priority = 3
          JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
          mailto=semail'�����ʼ���ַ
          JMail.AddRecipient(mailto)
          JMail.HTMLBody = book
          JMail.Body = "���"
          on error resume next
          JMail.Send(stupMailServer) 'SMTP��������ַ
		   
		rs.movenext
		loop
		end if
		JMail.Close()
        Set JMail = Nothing

end if	

end if
%>
<SCRIPT LANGUAGE="JavaScript">
<!--//
function checkmessage()
{   
    if (document.form.messagename.value.length<1)
	{
        alert("����д�����ˣ�");
        document.form.messagename.focus();
        return false;
    }

	
    if (document.form.messagetext.value.length<1)
	{
        alert("����д���ݣ�");
        document.form.messagetext.focus();
        return false;
    }
	    if (document.form.messagetext.value.length>1000)
	{
        alert("�������������1000���ڣ�");
        document.form.messagetext.focus();
        return false;
    }
}
//-->
</SCRIPT>
<form name="form" method="post" action="" onSubmit="return checkmessage()">
      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <tr>
    <td height="20" bgcolor="#799AE1" align="center">
      <table width="98%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td align="center"><font color="#FFFFFF" style="font-size:14px">�� �� �� ��</font></td>
          <td width="35">�� </td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"> <br>
        <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#D6DFF7">
          <tr bgcolor="#FFFFFF">
            <td width="27%" height="26" align="right">�����ˣ�</td>
            <td><input name="messagename" type="text" id="messagename" size="20" maxlength="20" value="<%=session("adminname")%>"></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td height="26" align="right">�����ˣ�</td>
            <td>
            
            <select name="ggwei1" id="ggwei1" size="1" style="position:absolute; left: 310px; top: 81px; width:120px; height:20px; clip: rect(0 120 21 100)" onChange="ggwei.value=ggwei1.value;ggwei.select()">
             <option >��ѡ������룡</option>
            <option value="all">ȫ����Ա</option>
		<%
		
		set rsggwei=conn.execute("select * from "&jieducm&"_register where dst<>0")
		if rsggwei.bof and rsggwei.eof then
		else
		do while not rsggwei.eof
		response.write "<option value="""&rsggwei("username")&""" >"&rsggwei("username")&"</option>"
		rsggwei.movenext
		loop
		end if
		rsggwei.close
		set rsggwei=nothing
		%>
              </select>
       <input type="text" style="position:absolute; left:311px; top:81px; width:102px; height:21px;" name="ggwei" value="<%
		if request("action")="xg" then
		call ggweitile(rsxg("HL_ggwei"),"HL_title")
		else
		response.write "��ѡ������룡" 
		end if
		%>" id="ggwei" onFocus="if(value=='��ѡ������룡') {value=''}" onBlur="if 
(value=='') {value='��ѡ������룡'}">
        </div>            </td>
          </tr>
          
          <tr bgcolor="#FFFFFF">
            <td align="right" valign="top">�������ݣ�</td>
            <td><textarea name="messagetext" cols="39" rows="5"></textarea>            </td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td height="30">��</td>
            <td><input type="submit" name="Submit" value="������Ϣ">
&nbsp;
<input type="reset" value="ȡ������" name="reset"></td>
          </tr>
        </table>
        <br>
    </td>
  </tr>
</table></form>                                               
</body>                                                                        
</html>     