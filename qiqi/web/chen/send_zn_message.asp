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
messagelxfs=HtmlEncode(trim(request.form("messagelxfs")))
if messagelxfs<>"" then
messagetext=HtmlEncode(trim(request.form("messagetext")))
messagename=HtmlEncode(trim(request.form("messagename")))
uid=HtmlEncode(trim(request.form("ggwei")))
if uid<>"all" then
 	    Set rs=server.createobject("ADODB.RECORDSET")
        rs.open "Select * From "&jieducm&"_register where username='"&uid&"'" ,Conn,3,3 
		if rs.eof then
		   response.write("<script>alert('�Բ���,�޴��û�');history.back();</script>")
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
rs2("messagename")=messagename'������
rs2("messagelxfs")=messagelxfs'����
rs2("messagetext")=messagetext'����
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
		rs("class")="����վ����Ϣ"
		rs("content")="����Ա��"&uid&"������վ����Ϣ"
		rs("jiegou")="�ɹ�"
    	rs.update
    	rs.close
		
    Response.Write("<script>alert('���ͳɹ�!');window.location.href='zn_message.asp';</script>")
	response.End()		
                                                             
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
          <td align="center"><font color="#FFFFFF" style="font-size:14px">�� �� վ �� �� Ϣ</font></td>
          <td width="35">��</td>
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
		response.write "��ѡ������룡" 
		end if
		%>" id="ggwei" onFocus="if(value=='��ѡ������룡') {value=''}" onBlur="if 
(value=='') {value='��ѡ������룡'}">
        </div>
            
            </td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td height="26" align="right">��&nbsp; �⣺</td>
            <td><input name="messagelxfs" type="text" id="messagelxfs2" size="40" maxlength="50"></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td align="right" valign="top">��&nbsp; �ݣ�</td>
            <td><textarea name="messagetext" cols="39" rows="5"></textarea>
              <input name="id" type="hidden" value="<%=request("id")%>"></td>
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
<%
set rs=nothing
conn.close
set conn=nothing%>                                                
</body>                                                                        
</html>     