<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
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
option explicit
response.buffer=true	
Const PurviewLevel=0    '����Ȩ��
Const CheckChannelID=0    '����Ƶ����0Ϊ���������Ƶ��
%>
<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim ChannelID,rs,sql
dim FoundErr,ErrMsg
dim AdminPurview_Channel

ChannelID=trim(request("ChannelID"))
if ChannelID="" then
	ChannelID=0
else
	ChannelID=Clng(ChannelID)
end if
%>
<html>
<head>
<title>�鿴����Ȩ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="title"><strong>�� �� �� �� Ȩ ��</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>��������</strong></td>
    <td height="30"><a href="Admin_ShowPurview.asp">���й���Ȩ��</a> | <a href="Admin_ShowPurview.asp?ChannelID=2">����Ƶ��Ȩ��</a> | <a href="Admin_ShowPurview.asp?ChannelID=3">���Ƶ��Ȩ��</a> | <a href="Admin_ShowPurview.asp?ChannelID=4">ͼƬƵ��Ȩ��</a></td>
  </tr>
</table>
<%
response.write "<br>�����ڵ�λ�ã��鿴����Ȩ��&nbsp;&gt;&gt;&nbsp;<font color='#ff6600'>"
select case ChannelID
	case 0
		response.write "���й���Ȩ��"
	case 2
		response.write "����Ƶ��Ȩ��"
	case 3
		response.write "���Ƶ��Ȩ��"
	case 4
		response.write "ͼƬƵ��Ȩ��"
	case else
		response.write "����Ĳ���"
end select
response.write "</font><br> "
if ChannelID=0 then
	call ShowAllPurview()
else
	call ShowChannelPurview()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
'call CloseConn()
%>
</body>
</html>

<%
sub ShowAllPurview()
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td colspan="2" class="title"><strong>�� �� Ƶ ��</strong> 
      <%
		if AdminPurview_Article=1 then response.write "��Ƶ������Ա��"
		if AdminPurview_Article=2 then response.write "����Ŀ�ܱࣩ"
		if AdminPurview_Article=3 then response.write "����Ŀ����Ա��"
		if AdminPurview_Article=4 then response.write "����Ȩ�ޣ�"
		%>
    </td>
    <td height="22" colspan="2"><strong>����Ƶ��</strong> 
      <%
		if AdminPurview_Soft=1 then response.write "��Ƶ������Ա��"
		if AdminPurview_Soft=2 then response.write "����Ŀ�ܱࣩ"
		if AdminPurview_Soft=3 then response.write "����Ŀ����Ա��"
		if AdminPurview_Soft=4 then response.write "����Ȩ�ޣ�"
		%>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
    <td width="30%">������ӡ�������Ŀ��ר��</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Article=1 then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
    <td width="30%">������ӡ�������Ŀ</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Soft=1 then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
    <td width="30%">���Թ����������ۼ����»���վ</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Article=1 then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
    <td width="30%">���Թ����������</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Soft=1 then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
    <td width="30%">���Թ���ר������</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Article<=2 then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
    <td width="30%">���Թ����������վ</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Soft=1 then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
    <td width="30%">����Ŀ����¼�롢��ˡ�����Ȩ��</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Article<=2 then
			response.write "<font color='#0066cc'>ȫ��Ȩ��</font>"
		elseif AdminPurview_Article=3 then
			response.write "<a href='Admin_ShowPurview.asp?ChannelID=2'><font color='#0066cc'>����Ȩ��</font></a>"
		else
			response.write "<font color='#ff6600'>��Ȩ��</font>"
		end if
	  %>
    </td>
    <td width="30%">����Ŀ���¼�롢��ˡ�����Ȩ��</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Soft<=2 then
			response.write "<font color='#0066cc'>ȫ��Ȩ��</font>"
		elseif AdminPurview_Soft=3 then
			response.write "<a href='Admin_ShowPurview.asp?ChannelID=3'><font color='#0066cc'>����Ȩ��</font></a>"
		else
			response.write "<font color='#ff6600'>��Ȩ��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="title"> 
    <td colspan="2"><strong>ͼƬƵ��</strong> 
      <%
		if AdminPurview_Photo=1 then response.write "��Ƶ������Ա��"
		if AdminPurview_Photo=2 then response.write "����Ŀ�ܱࣩ"
		if AdminPurview_Photo=3 then response.write "����Ŀ����Ա��"
		if AdminPurview_Photo=4 then response.write "����Ȩ�ޣ�"
		%>
    </td>
    <td height="22" colspan="2"><strong>����Ƶ��</strong></td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
    <td width="30%">������ӡ�������Ŀ</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Photo=1 then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
    <td width="30%">�ظ����� </td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Guest,"Reply")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
<td width="30%">���Թ����������</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Photo=1 then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
    <td width="30%">�޸����� </td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Guest,"Modify")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
<td width="30%">���Թ����������վ</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Photo=1 then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
    <td width="30%">ɾ������</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Guest,"Del")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
<td width="30%">����ĿͼƬ¼�롢��ˡ�����Ȩ��</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Photo<=2 then
			response.write "<font color='#0066cc'>ȫ��Ȩ��</font>"
		elseif AdminPurview_Photo=3 then
			response.write "<a href='Admin_ShowPurview.asp?ChannelID=4'><font color='#0066cc'>����Ȩ��</font></a>"
		else
			response.write "<font color='#ff6600'>��Ȩ��</font>"
		end if
	  %>
    </td>
    <td width="30%">�������</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Guest,"Check")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="title"> 
    <td colspan="4" height="22"><strong>��վ����Ȩ��</strong><strong> </strong></td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
<td width="30%">�����޸��Լ��ĵ�¼����</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"ModifyPwd")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
    <td width="30%">�����޸���վ��Ϣ����  
    </td>
    <td align="center" width="20%"> 
      <%if AdminPurview=1 then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
<td width="30%">���Խ���Ƶ������</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"Channel")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
    <td width="30%">���Թ�����վ���</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"AD")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
<td width="30%">���Թ�����վ����</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"Announce")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
    <td width="30%">���Թ�����������</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"FriendSite")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
<td width="30%">���Թ�����վ����</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"Vote")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
    <td width="30%">���Թ�����վͳ��</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"Counter")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
<td width="30%">���Թ���ע���û�</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"User")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
    <td width="30%">���Թ����ʼ��б�</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"MailList")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
<td width="30%">���Թ�����ɫģ��</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"Skin")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
    <td width="30%">���Թ���������ģ��</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"Layout")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
<td width="30%">���Թ���JS����</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"JS")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
    <td width="30%">���Թ����ϴ��ļ�</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"UpFile")=True then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
<td width="30%">���Թ������ݿ�</td>
    <td align="center" width="20%"> 
      <%if AdminPurview=1 then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
    <td width="30%">&nbsp;</td>
    <td align="center" width="20%">&nbsp; </td>
  </tr>
</table>
<%
end sub


Sub ShowChannelPurview()
	dim AdminChannel_Name
	select case ChannelID
	case 2
		AdminPurview_Channel=AdminPurview_Article
		AdminChannel_Name="����"
	case 3
		AdminPurview_Channel=AdminPurview_Soft
		AdminChannel_Name="���"
	case 4
		AdminPurview_Channel=AdminPurview_Photo
		AdminChannel_Name="ͼƬ"
	end select

%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="title"> 
    <td colspan="2" height="22"><strong><%=AdminChannel_Name%>Ƶ��</strong> 
      <%
		if AdminPurview_Channel=1 then response.write "��Ƶ������Ա��"
		if AdminPurview_Channel=2 then response.write "����Ŀ�ܱࣩ"
		if AdminPurview_Channel=3 then response.write "����Ŀ����Ա��"
		if AdminPurview_Channel=4 then response.write "����Ȩ�ޣ�"
		%>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
<td width="30%">������ӡ�������Ŀ</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Channel=1 then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
<td width="30%">���Թ���<%=AdminChannel_Name%>����</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Channel=1 then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
<td width="30%">���Թ���<%=AdminChannel_Name%>����վ</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Channel=1 then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <%if  ChannelID=2 then%>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
<td width="30%">���Թ���ר������</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Channel<=2 then
			response.write "<font color='#0066cc'>��</font>"
		else
			response.write "<font color='#ff6600'>��</font>"
		end if
	  %>
    </td>
  </tr>
  <%end if%>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
<td width="30%">����Ŀ<%=AdminChannel_Name%>¼�롢��ˡ�����Ȩ��</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Channel<=2 then
			response.write "<font color='#0066cc'>ȫ��Ȩ��</font>"
		elseif AdminPurview_Channel=3 then
			response.write "<font color='#0066cc'>����Ȩ��</font>"
		else
			response.write "<font color='#ff6600'>��Ȩ��</font>"
		end if
	  %>
    </td>
  </tr>
</table>
 
<% 
	if AdminPurview_Channel=3 then
		dim arrShowLine(10)
		for i=0 to ubound(arrShowLine)
			arrShowLine(i)=False
		next
		dim sqlClass,rsClass,i,iDepth
		select case ChannelID
		case 2
			sqlClass="select * From ArticleClass order by RootID,OrderID"
		case 3
			sqlClass="select * From SoftClass order by RootID,OrderID"
		case 4
			sqlClass="select * From PhotoClass order by RootID,OrderID"
		end select
		set rsClass=server.CreateObject("adodb.recordset")
		rsClass.open sqlClass,conn,1,1
	%>
	<table width="100%" border="0" cellspacing="1" cellpadding="0" class="border">
	  <tr align="center" class="title">
		<td height="22"><strong>��Ŀ����</strong></td>
		<td width="100" height="22"><strong>¼��</strong></td>
		<td width="100"><strong>���</strong></td>
		<td width="100" height="22"><strong>����</strong></td>
	  </tr>
		<% 
		do while not rsClass.eof 
		%>
	  <tr class="tdbg">
		<td><% 
		iDepth=rsClass("Depth")
		if rsClass("NextID")>0 then
			arrShowLine(iDepth)=True
		else
			arrShowLine(iDepth)=False
		end if
		if iDepth>0 then
			for i=1 to iDepth 
				if i=iDepth then 
					if rsClass("NextID")>0 then 
						response.write "<img src='images/tree_line1.gif' width='17' height='16' valign='abvmiddle'>" 
					else 
						response.write "<img src='images/tree_line2.gif' width='17' height='16' valign='abvmiddle'>" 
					end if 
				else 
					if arrShowLine(i)=True then
						response.write "<img src='images/tree_line3.gif' width='17' height='16' valign='abvmiddle'>" 
					else
						response.write "<img src='images/tree_line4.gif' width='17' height='16' valign='abvmiddle'>" 
					end if
				end if 
			next 
		  end if 
		  if rsClass("Child")>0 then 
			response.write "<img src='Images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>" 
		  else 
			response.write "<img src='Images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>" 
		  end if 
		  if rsClass("Depth")=0 then 
			response.write "<b>" 
		  end if 
		  response.write rsClass("ClassName")
		  %>
		</td>
		<td align="center"><%
		if AdminPurview_Channel=3 then
			if CheckClassMaster(rsClass("ClassInputer"),AdminName)=True then
				response.write "<font color='#0066cc'>��</font>"
			else
				response.write "<font color='#ff6600'>��</font>"
			end if
		end if
		%>
		</td>
		<td align="center"><%
		if AdminPurview_Channel=3 then
			if CheckClassMaster(rsClass("ClassChecker"),AdminName)=True then
				response.write "<font color='#0066cc'>��</font>"
			else
				response.write "<font color='#ff6600'>��</font>"
			end if
		end if
		%></td>
		<td align="center"><%
		if AdminPurview_Channel=3 then
			if CheckClassMaster(rsClass("ClassMaster"),AdminName)=True then
				response.write "<font color='#0066cc'>��</font>"
			else
				response.write "<font color='#ff6600'>��</font>"
			end if
		end if
		%></td>
	  </tr>
		<% 
		rsClass.movenext 
		loop 
		rsClass.close
		set rsClass=nothing
		%>
	</table>
	<%
	end if
end sub
%>
