<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">

<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
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
action=request("action")
if action="fa" then
uid=HtmlEncode(trim(request.form("uid")))
content=(trim(request.form("content")))

  Set rs=server.createobject("ADODB.RECORDSET")
  rs.open "Select * From "&jieducm&"_toushu  where id="&uid&"" ,Conn,3,3  
		rs("bianjie")=content
    	rs.update
    	rs.close
		
		
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="���"
		rs("content")=session("useridname")&"���"&name1&"�ɹ�"
		rs("jiegou")="�ɹ�"
    	rs.update
    	rs.close	
		
		Response.Write("<script>alert('���ɹ�!');window.location.href='complaintmy.asp';</script>")

end if		
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<style type="text/css">
input{border:#e1af3c 1px solid;}
</style>

</head>

<body>
<!--#include file=../inc/jieducm_top.asp-->
<table width="910" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="margin6">
  <tr>
    <td width="147" valign="top"><!--#include file="userleft.asp"-->
 
    </td><td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="images/shoucangqu.gif">
      <tr>
        <td width="19%" align="center" class="font12b">��Ҫ���</td>
        <td width="81%" height="30">&nbsp;</td>
      </tr>
    </table>
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
       <FORM name=formf method=post action="?action=fa"  onSubmit="return save_onclick12()">
      <table width="100%" border="0" cellpadding="0" cellspacing="0" class="bordera2">
        <tr>
          <td height="300" valign="middle"><table width="700" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="110" height="40" align="center">������ԭ��</td>
              <td width="490"><%=rsm("content")%></td>
            </tr>
            <tr>
              <td height="10" align="center">�Է���������</td>
              <td><%=rsm("bianjie2")%></td>
            </tr>
            <tr>
              <td height="40" align="center">��Ҫ�������ݣ�</td>
              <td>
               <textarea name="Content" id="content" style="display:none"></textarea> 
              <iframe id="Editor" name="Editor" src="../HtmlEditor/index.html?ID=content" frameborder="0" marginheight="0"     marginwidth="0" scrolling="No" style="height:320px;width:100%"></iframe></td>
            </tr>
            <tr>
              <td height="70">&nbsp;</td>
              <td valign="middle">&nbsp;&nbsp;&nbsp;
                <input type="submit" name="button" id="button" value="�ύ">
                <input type="reset" name="button2" id="button2" value="����">
                <input name="uid" type="hidden" id="uid" value="<%=rsm("id")%>" />
              <input name="id" type="hidden" id="id" value="<%=id%>" />
              </td>
            </tr>
          </table></td>
        </tr>
      </table>
      </FORM>
      <% end if%>
    </td>
  </tr>
</table>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
 call CloseConn()
%>
</body>
</html>
<SCRIPT language=javascript>
function  save_onclick12()
{

  
 var content=document.formf.content.value;
  if(content=="")
  {
    alert("���ݲ���Ϊ�գ�");
	
	return false;
	}
	
  return true;  
}
</SCRIPT>