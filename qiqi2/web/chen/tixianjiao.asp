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
id =request("id")
Sql3 = "select * from "&jieducm&"_tixian where id="&id&""
Set Rs3 = Server.CreateObject("Adodb.RecordSet")
Rs3.Open Sql3,conn,3,3
if not rs3.eof then
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�������--����֧�������׺�</title>
   <SCRIPT language=javascript>
function  save_onclick22()
{

    var Price=document.form1.zfbnum.value;
  if(Price=="")
  {
    alert("������֧�������׺ţ�");
	document.form1.zfbnum.focus();
	return false;
	}
  

  return true;  
}
</script>  
</head>

<body>
<form id="form1" name="form1" method="post" action="tixianok.asp?id=<%=rs3("id")%>&action=ok" onSubmit="return save_onclick22();">
  <table width="336" border="0">
    <tr>
      <td width="125">������ˮ�ţ�</td>
      <td width="201"><%=rs3("pid")%></td>
    </tr>
    <tr>
      <td>�����û�����</td>
      <td><%=rs3("username")%></td>
    </tr>
    <tr>
      <td>����֧�����ţ�</td>
      <td><%=rs3("ReZfb")%></td>
    </tr>
    <tr>
      <td>��ʵ������</td>
      <td><%=rs3("ReRName")%></td>
    </tr>
    <tr>
      <td>���׺ţ�</td>
      <td><input type="text" name="zfbnum" id="zfbnum" /></td>
    </tr>
    <tr>
      <td><input type="submit" name="button" id="button" value="�ύ" /></td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
<% end if%>