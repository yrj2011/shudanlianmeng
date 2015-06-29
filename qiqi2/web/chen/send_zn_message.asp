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
messagelxfs=HtmlEncode(trim(request.form("messagelxfs")))
if messagelxfs<>"" then
messagetext=HtmlEncode(trim(request.form("messagetext")))
messagename=HtmlEncode(trim(request.form("messagename")))
uid=HtmlEncode(trim(request.form("ggwei")))
if uid<>"all" then
 	    Set rs=server.createobject("ADODB.RECORDSET")
        rs.open "Select * From "&jieducm&"_register where username='"&uid&"'" ,Conn,3,3 
		if rs.eof then
		   response.write("<script>alert('对不起,无此用户');history.back();</script>")
		   response.End()
		end if
end if	
	
ip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
If ip = "" Then 
ip = Request.ServerVariables("REMOTE_ADDR") 
end if
sql="select * from "&jieducm&"_china_message"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,3
rs2.addnew
rs2("messagename")=messagename'发送人
rs2("messagelxfs")=messagelxfs'标题
rs2("messagetext")=messagetext'内容
rs2("uid")=uid
rs2("ip")=ip
rs2("zn")="yes"
rs2("messagedate")=now
rs2("hits")=0
rs2.update
rs2.close  
set rs2=nothing 

 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
		rs("class")="发送站内信息"
		rs("content")="管理员给"&uid&"发送了站内信息"
		rs("jiegou")="成功"
    	rs.update
    	rs.close
		
    Response.Write("<script>alert('发送成功!');window.location.href='zn_message.asp';</script>")
	response.End()		
                                                             
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
          <td align="center"><font color="#FFFFFF" style="font-size:14px">发 送 站 内 消 息</font></td>
          <td width="35">　</td>
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
            <td height="26" align="right">标&nbsp; 题：</td>
            <td><input name="messagelxfs" type="text" id="messagelxfs2" size="40" maxlength="50"></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td align="right" valign="top">内&nbsp; 容：</td>
            <td><textarea name="messagetext" cols="39" rows="5"></textarea>
              <input name="id" type="hidden" value="<%=request("id")%>"></td>
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
set rs=nothing
conn.close
set conn=nothing%>                                                
</body>                                                                        
</html>     