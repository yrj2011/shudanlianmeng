<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->

<%

action=request("action")
if action="ok" then
uid=request("uid")
content=request("content")

  Set rs=server.createobject("ADODB.RECORDSET")
  rs.open "Select * From "&jieducm&"_toushu  where id="&uid&"" ,Conn,3,3  
		rs("guang")=content
    	rs.update
    	rs.close
			Response.Write("<script>alert('�����ɹ�!');window.location.href='shenshu.asp';</script>")
end if		

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" href="Style.css">
<title>���ߴ�ý</title>
</head>

<body>
    <%
	id=request("id")

	Sql = "select * from "&jieducm&"_toushu where id="&id&""
	Set Rsm = Server.CreateObject("Adodb.RecordSet")
	Rsm.Open Sql,conn,1,1
	IF Rsm.Eof Then
	   Response.Write("<script>alert('�������ô���!');history.back();</script>")
	   response.end()
	else
	%>
     <form name="myform" method="Post" action="?action=ok" >
<table width="700" border="0" align="center" cellpadding="0" cellspacing="0" class="border" >
            <tr>
              <td height="40" align="center">&nbsp;ԭ��</td>
              <td><%=rsm("class")%></td>
            </tr>
            <tr>
              <td height="40" align="center">���⣺</td>
              <td><%=rsm("title")%></td>
            </tr>
            <tr>
              <td width="110" height="40" align="center">������ԭ��</td>
              <td width="490"><%=rsm("content")%></td>
            </tr>
            <tr>
              <td height="10" align="center">�Է���������</td>
              <td><textarea name="content2" id="content2" cols="55" rows="6" readonly="readonly"><%=rsm("bianjie2")%></textarea></td>
            </tr>
            <tr>
              <td height="40" align="center">���߷��������ݣ�</td>
        <td>
              
                  <textarea name="content3" id="content3" cols="55" rows="6" readonly="readonly"><%=rsm("bianjie")%></textarea>
</p></td>
            </tr>
           
            <tr>
              <td height="70"><div align="center">�ٷ����</div></td>
              <td valign="middle"><textarea name="content" id="content" cols="55" rows="6" ><%=rsm("guang")%></textarea>               </td>
            </tr> <tr>
              <td height="70">&nbsp;</td>
              <td valign="middle"><input type="submit" name="button" id="button" value="��  ��"   />
              <input name="uid" type="hidden" id="uid" value="<%=rsm("id")%>" />
               <input name="id" type="hidden" id="id" value="<%=rsm("id")%>" />
               <input type="button" name="button2" id="button2" value="��  ��" onClick="javascript:history.back();" /></td>
            </tr>
          </table>
     </form>
          <% end if%>
		  <%conn.close
		  set conn=nothing%>
</body>
</html>
