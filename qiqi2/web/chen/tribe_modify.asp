
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
dim rs, sql, strPurview,ok,id,fid,qq,num,iCount,rsm
dim Action,rsr,FoundErr,ErrMsg,tribename,pic,tribeinfo,startid,triberad
id=request("id")
action=request("action")
	
if action="modify" then
	 Sql= "select * from "&jieducm&"_tribe where id="&request("startid")&""
	 Set rs=server.createobject("ADODB.RECORDSET")
	 Rs.Open Sql,conn,3,3
	 if not rs.eof then
		rs("tribename")=request.Form("tribename")	
    	rs("pic")=request.Form("pic")
		rs("tribeinfo")=Request.Form("tribeinfo")
		rs("triberad")=request.Form("triberad")
    	rs.update
		rs.close
	 end if
	 Response.Write("<script>alert('�޸ĳɹ�!');window.location.href='tribemanage.asp';</script>")
	 response.End()
end if
%>
<html>
<head>
<title>����Ա����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
<SCRIPT language=javascript>
function  save_onclick()
{

 var Price=document.aspnetForm.tribename.value;
  if(Price=="")
  {
    alert("����������Ϊ�գ�");
	document.aspnetForm.tribename.focus();
	return false;
	}
	
 var ProUrl=document.aspnetForm.pic.value;
  if(ProUrl=="")
  {
    alert("������²���Ϊ�գ�");
	document.aspnetForm.pic.focus();
	return false;
	}
	 var triberad=document.aspnetForm.triberad.value;
  if(triberad=="")
  {
    alert("���伤���벻��Ϊ�գ�");
	document.aspnetForm.triberad.focus();
	return false;
	}
	
 var ProUrl=document.aspnetForm.tribeinfo.value;
  if(ProUrl=="")
  {
    alert("�����鲻��Ϊ�գ�");
	document.aspnetForm.tribeinfo.focus();
	return false;
	}
  return true;  
}
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="title"><strong>�������</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>��������</strong></td>
    <td height="30"><a href="tribe_add.asp">��Ӳ���  </a> <a  href="tribemanage.asp">�������</a></td>
  </tr>
</table>
<%
   sql = "select * from "&jieducm&"_tribe where id="&id&""
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then
				 tribename=rs("tribename")
				 pic=rs("pic")
				 tribeinfo=rs("tribeinfo")
				 startid=rs("id")
				 triberad=rs("triberad")
				 username=rs("username")
				 rs.close
				end if
%>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
 <form action="?action=modify" method="POST" name="aspnetForm" id="aspnetForm" onSubmit="return save_onclick()">

<input name="startid" type="hidden" value="<%=startid%>">
	<tr>
	  <td colspan="3" align="center" class="title">&nbsp;</td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">�������ƣ�</DIV></td>
      <td class="tdbg"><label>
        <input type="text" name="tribename" value="<%=tribename%>">
      </label></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">������£�</DIV></td>
      <td class="tdbg"><label>
        <input type="text" name="pic" value="<%=pic%>"> 
       &nbsp; 
     <input type="button" name="Submit" value="�ϴ�ͼƬ"  onClick="window.open('../upload_flash.asp?formname=aspnetForm&editname=pic&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
      </label></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">�����룺</DIV></td>
      <td class="tdbg"><label>
			<input type="text" name="triberad" value="<%=triberad%>">
      </label></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">������</DIV></td>
      <td class="tdbg"><label>
        <input type="text" name="username" value="<%=username%>" disabled="disabled">
      ���������޸�</label></td>
    </tr>
    <tr class="tdbg">
      <td width="35%" class="tdbg"><DIV align="right">�����飺</DIV></td>
      <td width="65%" class="tdbg"><label>
        <textarea name="tribeinfo" cols="40" rows="5"><%=tribeinfo%></textarea>
      </label></td>
    </tr>
    
    
    <tr class="tdbg">
      <td colspan="2"><table id="PurviewDetail" width="100%" border="0" cellspacing="10" cellpadding="0" style="display:none">
        <tr>
          <td colspan="2" align="center" class="title"><strong>�� �� Ա Ȩ �� �� ϸ �� ��</strong></td>
        </tr>

   
      </table></td>
    </tr>
    <tr> 
      <td height="40" colspan="2" align="center" class="tdbg"><input  type="submit" name="Submit" value=" ��&nbsp;&nbsp;�� " style="cursor: hand;background-color: #cccccc;">&nbsp; 
		<input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='tribemanage.asp'" style="cursor: hand;background-color: #cccccc;"></td>
    </tr>
  </form>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" align="center" background=images/admin_bg_bottom.gif><font color=#cccccc>Copyright &copy; 2013</font> <a href='http://www.778892.com' target="_blank"><font color=#ff6600><strong>�������ƽ̨</strong></font></a> <font color=#cccccc>. All Rights Reserved.</font></td>
  </tr>
</table>
</body>
</html>
