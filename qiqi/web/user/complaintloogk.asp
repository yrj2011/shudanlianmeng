<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">

<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#include file="../background.asp"-->
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
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<link href="../css/fenye.css" rel="stylesheet" type="text/css" />
<style type="text/css">
input{border:#e1af3c 1px solid;}
</style>

</head>

<body>
<!--#include file=../inc/jieducm_top.asp-->
<table width="910" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="margin6">
  <tr>
    <td width="147" valign="top"><!--#include file="userleft.asp"-->
  
    </td><td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../images/shoucangqu.gif">
      <tr>
        <td width="19%" align="center" class="font12b">�������</td>
        <td width="81%" height="30">&nbsp;</td>
      </tr>
    </table>
    <%
	id=request("id")
	my=request("my")
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
              <td height="40" align="center">����ԭ��</td>
              <td><%=rsm("class")%></td>
            </tr>
            <tr>
              <td width="110" height="40" align="center"><%if my="1" then%>
              ���������ݣ�
                <%elseif my="2" then%>
                �Է��������ݣ�
                <% end if%></td>
              <td width="490"><%=rsm("content")%></td>
            </tr>
            <tr>
              <td height="10" align="center"><%if my="1" then%>�Է���������<%elseif my="2" then%>�ұ�������<%end if%></td>
              <td><%=rsm("bianjie2")%></td>
            </tr>
            <tr>
              <td height="40" align="center"><%if my="1" then%>�ұ������ݣ�<%elseif my="2" then%>�Է��������ݣ�<% end if%></td>
              <td>
              
                 <%=rsm("bianjie")%>
</p></td>
            </tr>
           
            <tr>
              <td height="70"><div align="center">�ٷ����</div></td>
              <td valign="middle"><%=rsm("guang")%>               </td>
            </tr> <tr>
              <td height="70">&nbsp;</td>
              <td valign="middle"><input type="button" name="button" id="button" value="��    ��"  onclick="javascript:history.back();" /></td>
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
call  CloseConn()
%>
</body>
</html>
