
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷�����  �� �� ý
'�ٷ���ҳ��http://www.778892.com
'����֧�֣�xsh
'QQ;859680966
'�绰��15889679112
'��Ȩ����Ȩ�������ߴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2013 ���ߴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
option explicit
response.buffer=true	
Const PurviewLevel=1
'response.write "�˹��ܱ�WEBBOY��ʱ��ֹ�ˣ�"
'response.end
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/md5.asp"-->
<%
dim  sql, strPurview,ok,id,fid,qq,num,iCount,uid
dim Action,rs,FoundErr,ErrMsg
id=request("id")
uid=request("uid")
ok=request("ok")
if ok="add" then
	Sql = "select * from "&jieducm&"_register where id="&uid&""
	Set rs = Server.CreateObject("Adodb.RecordSet")
	rs.Open Sql,conn,1,1
	IF rs.Eof Then
	   Response.Write("<script>alert('�Բ����޴��û�!');window.history.back();</script>")
	   response.end()
	 else
	 username=rs("username")
	end if
	
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_keeper where keeper='"&request.form("qq")&"'" ,Conn,3,3  
		if not rs.eof then
		  Response.Write("<script>alert('�˺��ѱ������û���!');window.history.back();</script>")
	      response.end()
		else
	    rs.addnew
		rs("username")=username
		rs("keeper")=request.form("qq")
		rs("class")=request.form("classid")
		rs("now")=now()
    	rs.update
    	rs.close
		end if
 Response.Write("<script>alert('�󶨳ɹ�!');window.location.href='usermanage.asp';</script>")
 response.End()
end if
%>
<html>
<head>
<title>����Ա����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="title"><strong>���ƹ�</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>��������</strong></td>
    <td height="30">&nbsp;</td>
  </tr>
</table>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form method="post" action="?ok=add" name="form1" onSubmit="javascript:return CheckAdd();">
	<tr>
	  <td colspan="3" align="center" class="title">&nbsp;</td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><strong>���ࣺ</strong></td>
      <td class="tdbg"><select name="classid">
          <option value="1">�Ա���</option>
          <option value="2">������</option>
          <option value="3">�а���</option>
        </select></td>
    </tr>
    <tr class="tdbg">
      <td width="35%" class="tdbg"><strong>���ƹ�</strong></td>
      <td width="65%" class="tdbg"><input name="qq" type="text"></td>
    </tr>
    
    
    <tr class="tdbg">
      <td colspan="2"><table id="PurviewDetail" width="100%" border="0" cellspacing="10" cellpadding="0" style="display:none">
        <tr>
          <td colspan="2" align="center" class="title"><strong>�� �� Ա Ȩ �� �� ϸ �� ��</strong></td>
        </tr>

   
      </table></td>
    </tr>
    <tr> 
      <td height="40" colspan="2" align="center" class="tdbg"><input type="hidden" name="uid" value="<%=id%>">
      <input  type="submit" name="Submit" value=" ��&nbsp;&nbsp;�� " style="cursor: hand;background-color: #cccccc;">&nbsp; 
		<input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='usermanage.asp'" style="cursor: hand;background-color: #cccccc;"></td>
    </tr>
  </form>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" align="center" background=images/admin_bg_bottom.gif><font color=#cccccc>Copyright &copy; 2013</font> <a href='http://www.778892.com' target="_blank"><font color=#ff6600><strong>�����������ƽ̨</strong></font></a> <font color=#cccccc>. All Rights Reserved.</font></td>
  </tr>
</table>
</body>
</html>
