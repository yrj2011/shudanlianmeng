<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>发送站内消息</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>
<%
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
		   response.write("<script>alert('对不起,无此用户');history.back();</script>")
		   response.End()
		else
		  semail=rs("mail")
		  title=messagelxfs '邮件标题
          book=messagetext '邮件内容
          Set JMail = Server.CreateObject("JMail.Message") 
          JMail.Charset = "gb2312"
          JMail.From = stupMailDomain 'SMTP域名
          JMail.FromName = messagename '发件人
          JMail.Subject = title 
          JMail.MailServerUserName = stupMailServerUserName'用户名
          JMail.MailServerPassword = stupMailServerPassWord'密码
          JMail.Priority = 3
          JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
          mailto=semail'接收邮件地址
          JMail.AddRecipient(mailto)
          JMail.HTMLBody = book
          JMail.Body = "你好"
          on error resume next
          JMail.Send(stupMailServer) 'SMTP服务器地址
          JMail.Close()
          Set JMail = Nothing
		end if
	 Response.Write("<script>alert('发送成功!');window.location.href='send_ms_message.asp';</script>")
	response.End()	
else

        Set rs=server.createobject("ADODB.RECORDSET")
        rs.open "Select mail  From "&jieducm&"_register " ,Conn,3,3 
		if rs.eof then
		   response.write("<script>alert('对不起,操作有误！');history.back();</script>")
		   response.End()
		else
		do while not rs.eof
		  semail=rs("mail")
		  title=messagelxfs '邮件标题
          book=messagetext '邮件内容
          Set JMail = Server.CreateObject("JMail.Message") 
          JMail.Charset = "gb2312"
          JMail.From = stupMailDomain 'SMTP域名
          JMail.FromName = messagename '发件人
          JMail.Subject = title 
          JMail.MailServerUserName = stupMailServerUserName'用户名
          JMail.MailServerPassword = stupMailServerPassWord'密码
          JMail.Priority = 3
          JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
          mailto=semail'接收邮件地址
          JMail.AddRecipient(mailto)
          JMail.HTMLBody = book
          JMail.Body = "你好"
          on error resume next
          JMail.Send(stupMailServer) 'SMTP服务器地址
		rs.movenext
		loop
		end if
		JMail.Close()
        Set JMail = Nothing
	  Response.Write("<script>alert('发送成功!');window.location.href='send_ms_message.asp';</script>")
	response.End()	

end if	

end if
%>
<SCRIPT LANGUAGE="JavaScript">
<!--//
function checkmessage()
{   
    if (document.form.messagename.value.length<1)
	{
        alert("请填写发送人！");
        document.form.messagename.focus();
        return false;
    }
	    if (document.form.messagelxfs.value.length<1)
	{
        alert("请填写标题！");
        document.form.messagelxfs.focus();
        return false;
    }
	
    if (document.form.messagetext.value.length<1)
	{
        alert("请填写内容！");
        document.form.messagetext.focus();
        return false;
    }
	    if (document.form.messagetext.value.length>1000)
	{
        alert("请把字数控制在1000以内！");
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
          <td align="center"><font color="#FFFFFF" style="font-size:14px">发 送 邮 件</font></td>
          <td width="35">　 </td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"> <br>
        <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#D6DFF7">
          <tr bgcolor="#FFFFFF">
            <td width="27%" height="26" align="right">发送人：</td>
            <td><input name="messagename" type="text" id="messagename" size="20" maxlength="20" value="<%=session("adminname")%>"></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td height="26" align="right">接收人：</td>
            <td>
            
            <select name="ggwei1" id="ggwei1" size="1" style="position:absolute; left: 310px; top: 81px; width:120px; height:20px; clip: rect(0 120 21 100)" onChange="ggwei.value=ggwei1.value;ggwei.select()">
             <option >请选择或输入！</option>
            <option value="all">全部会员</option>
		<%
		
		set rsggwei=conn.execute("select * from "&jieducm&"_register")
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
		response.write "请选择或输入！" 
		end if
		%>" id="ggwei" onFocus="if(value=='请选择或输入！') {value=''}" onBlur="if 
(value=='') {value='请选择或输入！'}">
        </div>
            
            </td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td height="26" align="right">邮件标题：</td>
            <td><input name="messagelxfs" type="text" id="messagelxfs" size="40" maxlength="50"></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td align="right" valign="top">邮件标题内容：</td>
            <td><textarea name="messagetext" cols="39" rows="5"></textarea>
            </td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td height="30">　</td>
            <td><input type="submit" name="Submit" value="发送消息">
&nbsp;
<input type="reset" value="取消发送" name="reset"></td>
          </tr>
        </table>
        <br>
    </td>
  </tr>
</table></form>      
<%
conn.close
set conn=nothing%>                                         
</body>                                                                        
</html>     