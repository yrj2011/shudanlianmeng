<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="inc/conn.asp"-->
<!--#INCLUDE FILE="inc/md5.asp"-->
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
username=request("username")
action=request("action")
id=request("id")
if action="ok" then
       Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_keeper where id="&request("id")&""  ,Conn,3,3  
	  if  not rs.eof then	  
		rs("keeper")=request.form("username")
    	rs.update
    	rs.close
		response.Redirect("userbuyww.asp?username="&username&"")
		response.End()
	   end if
elseif action="del" then
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete from "&jieducm&"_keeper where id="&id&"",conn,3,3
response.Redirect("userbuyww.asp?username="&username&"")

end if
%>

<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
<style type="text/css">
<!--
.STYLE1 {font-size: 16px}
-->
</style>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">


<DIV style="WIDTH: 950px">
<DIV class=regheight 
style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 15px; PADDING-TOP: 15px"><BR>
  <BR>
  <DIV style="COLOR: red"><SPAN id=ctl00_AllBaseContentPlaceHolder_MsgLabel 
style="COLOR: red"></SPAN><%=username%>�󶨵��ƹ���</DIV>
  <%
			  	Sql = "Select * From "&jieducm&"_keeper where username='"&username&"'"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("���û����ް󶨵��ƹ�����")
				Else
				   i=1
					Do While Not Rs.Eof					
			  %>
 <FORM action="?action=ok&username=<%=username%>" method="POST" name="form<%=i%>" id="form<%=i%>" onSubmit="return save_onclick()">
  <table width="840" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="173"><div align="right">�󶨵�
	  <%if rs("class")=1 then
	  response.Write("�Ա��ƹ�����")
	  elseif rs("class")=2 then
	  response.Write("�����ƹ�����")
	  elseif rs("class")=3 then
	  response.Write("�а��ƹ�����")
	  end if
	  %>
	  </div></td>
     
      <td width="191"><input id="username" maxlength=50 name="username" value="<%=rs("keeper")%>"  ></td> 
      <td width="95">��ʱ�䣺</td>
      <td width="95"><%=rs("now")%></td>
      <td width="96"><input type="submit" name="button2" id="button2" value="�� ��">
&nbsp;
<input type="button" name="button" id="button" value="ɾ ��" onClick="window.location.href='userbuyww.asp?action=del&username=<%=username%>&id=<%=rs("id")%>'">
<input type="hidden" name="id" id="id" value="<%=rs("id")%>"></td>
    </tr>
  </table>
    </FORM>
  <%
                i=i+1
			  	Rs.MoveNext
				Loop
				End IF
			  %>
<DIV class=regsubmit>
  <div align="center"><br>
  &nbsp;  
  <INPUT name="��ť"  type=button style="CURSOR: pointer; HEIGHT: 30px" value="  ����  " onClick="history.back();">
  </div>
</DIV><BR><BR></DIV></DIV>


</DIV>
</BODY></HTML>
