
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
dim Action,rsr,FoundErr,ErrMsg,tribename,pic,tribeinfo,startid,triberad,tribeid
action=request("action")
id=request("id")	
if action="modify" then	  
	 Sql= "select * from "&jieducm&"_car where id="&request("car_id")&""
	 Set rs=server.createobject("ADODB.RECORDSET")
	 Rs.Open Sql,conn,3,3
	 if not rs.eof then
		rs("car_name")=request("car_name")
		rs("car_price")=request("car_price")
	    rs("car_fabudian")=request("car_fabudian")
    	rs("car_pic")=request("pic")
		rs("car_content")=request("car_content")
		rs("car_sort")=request("car_sort")
		rs("now")=now()
    	rs.update
		rs.close
	 else
	   Response.Write("<script>alert('�Բ����޴˵㿨!');history.back();</script>")
	   response.End()
	 end if
	 Response.Write("<script>alert('�޸ĳɹ�!');window.location.href='jieducm_car.asp';</script>")
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

 var car_name=document.aspnetForm.car_name.value;
  if(car_name=="")
  {
    alert("�㿨���Ʋ���Ϊ�գ�");
	document.aspnetForm.car_name.focus();
	return false;
	}
	
 var car_price=document.aspnetForm.car_price.value;
  if(car_price=="")
  {
    alert("�㿨�۸���Ϊ�գ�");
	document.aspnetForm.car_price.focus();
	return false;
	}
	 var car_fabudian=document.aspnetForm.car_fabudian.value;
  if(car_fabudian=="")
  {
    alert("�㿨�����㲻��Ϊ�գ�");
	document.aspnetForm.car_fabudian.focus();
	return false;
	}
	
var pic=document.aspnetForm.pic.value;
  if(pic=="")
  {
    alert("�㿨ͼƬ����Ϊ�գ�");
	document.aspnetForm.pic.focus();
	return false;
	}
	

  return true;  
}
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="title"><strong>��վ�㿨����</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>��������</strong></td>
    <td height="30"><a href="jieducm_caradd.asp">�����վ�㿨  </a> <a  href="jieducm_car.asp">��վ�㿨����</a></td>
  </tr>
</table>

<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
 <form action="?action=modify" method="POST" name="aspnetForm" id="aspnetForm" onSubmit="return save_onclick()">
<%
   sql = "select * from "&jieducm&"_car where id="&id&""
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then
%>
<input name="car_id" type="hidden" value="<%=rs("id")%>">
	<tr>

	  <td colspan="3" align="center" class="title">&nbsp;</td>
    </tr>
<tr class="tdbg">
      <td class="tdbg"><DIV align="right">����</DIV></td>
      <td class="tdbg"><label>
        <input name="car_sort" type="text" onKeyUp="if(isNaN(value))execCommand('undo')" size="5" value="<%=rs("car_sort")%>">
      ������ԽСԽ��ǰ��</label></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">�㿨���ƣ�</DIV></td>
      <td class="tdbg"><label>
        <input type="text" name="car_name"  value="<%=rs("car_name")%>">
      </label></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">�㿨�۸�</DIV></td>
      <td class="tdbg"><input type="text" name="car_price" onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=rs("car_price")%>" maxlength="5"></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">�㿨�Ͷ��ٷ����㣺</DIV></td>
      <td class="tdbg"><label>
        <input type="text" name="car_fabudian" value="<%=rs("car_fabudian")%>" onKeyUp="if(isNaN(value))execCommand('undo')" maxlength="5" >
      </label></td>
    </tr>
    <tr class="tdbg">
      <td class="tdbg"><DIV align="right">�㿨ͼƬ��</DIV></td>
      <td class="tdbg"><label>
        <input type="text" name="pic" id="pic" value="<%=rs("car_pic")%>">
      </label> &nbsp; 
     <input type="button" name="Submit" value="�ϴ�ͼƬ"  onClick="window.open('../upload_flash.asp?formname=aspnetForm&editname=pic&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
     ��ͼƬ���322*184��	 </td>
    </tr>
    
    <tr class="tdbg">
      <td width="35%" class="tdbg"><DIV align="right">�㿨��飺</DIV></td>
      <td width="65%" class="tdbg"><label>
        <textarea name="car_content" cols="40" rows="5"><%=rs("car_content")%></textarea>
      </label></td>
    </tr>
    
    

    <tr> 
      <td height="40" colspan="2" align="center" class="tdbg"><input  type="submit" name="Submit" value=" ��&nbsp;&nbsp;�� " style="cursor: hand;background-color: #cccccc;">&nbsp; 
		<input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='jieducm_car.asp'" style="cursor: hand;background-color: #cccccc;"></td>
    </tr>
<%end if%>
  </form>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" align="center" background=images/admin_bg_bottom.gif><font color=#cccccc>Copyright &copy; 2013</font> <a href='http://www.778892.com' target="_blank"><font color=#ff6600><strong>�����������ƽ̨</strong></font></a> <font color=#cccccc>. All Rights Reserved.</font></td>
  </tr>
</table>
</body>
</html>
