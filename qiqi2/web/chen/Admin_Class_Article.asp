<%@language=vbscript codepage=936 %>
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
Const PurviewLevel=2 
Const CheckChannelID=2
Const PurviewLevel_Article=1
%>
<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/admin_code_article.asp"-->
<%
dim Action,ParentID,SkinID,LayoutID,BrowsePurview,AddPurview,i,FoundErr,ErrMsg
dim SkinCount,LayoutCount
Action=trim(Request("Action"))
ParentID=trim(request("ParentID"))
if ParentID="" then
	ParentID=0
else
	ParentID=CLng(ParentID)
end if
%>
<html>
<head>
<title>������Ŀ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">

<style type="text/css">
<!--
.style1 {color: #FF6600}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr>
	<td colspan="2" align="center" class="title"><strong>�� �� �� Ŀ �� ��</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="80" height="30"><strong>��������</strong></td>
    <td height="30"><a href="Admin_Class_Article.asp">������Ŀ������ҳ</a> | <a href="Admin_Class_Article.asp?Action=Add">���������Ŀ</a>&nbsp;|&nbsp;<a href="Admin_Class_Article.asp?Action=Order">һ����Ŀ����</a>&nbsp;|&nbsp;<a href="Admin_Class_Article.asp?Action=OrderN">N����Ŀ����</a>&nbsp;|&nbsp;<a href="Admin_Class_Article.asp?Action=Reset">��λ����������Ŀ</a>&nbsp;|&nbsp;<a href="Admin_Class_Article.asp?Action=Unite">������Ŀ�ϲ�</a></td>
  </tr>
</table>
<%
if Action="Add" then
	call AddClass()
elseif Action="SaveAdd" then
	call SaveAdd()
elseif Action="Modify" then
	call Modify()
elseif Action="SaveModify" then
	call SaveModify()
elseif Action="Move" then
	call MoveClass()
elseif Action="SaveMove" then
	call SaveMove()
elseif Action="Del" then
	call DeleteClass()
elseif Action="Clear" then
	call ClearClass()
elseif Action="UpOrder" then 
	call UpOrder() 
elseif Action="DownOrder" then 
	call DownOrder() 
elseif Action="Order" then
	call Order()
elseif Action="UpOrderN" then 
	call UpOrderN() 
elseif Action="DownOrderN" then 
	call DownOrderN() 
elseif Action="OrderN" then
	call OrderN()
elseif Action="Reset" then
	call Reset()
elseif Action="SaveReset" then
	call SaveReset()
elseif Action="Unite" then
	call Unite()
elseif Action="SaveUnite" then
	call SaveUnite()
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
''call CloseConn() 'shiyu


sub main()
	dim arrShowLine(10)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	dim sqlClass,rsClass,i,iDepth
	sqlClass="select * From "&jieducm&"_ArticleClass order by RootID,OrderID"
	set rsClass=server.CreateObject("adodb.recordset")
	rsClass.open sqlClass,conn,1,1
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class="title"> 
    <td height="22" align="center"><strong>��Ŀ����</strong></td>
    
    <td width="80" align="center"><strong>��Ŀ����</strong></td>
    <td width="60" align="center"><strong>���Ȩ��</strong></td>
    <td width="300" height="22" align="center"><strong>����ѡ��</strong></td>
  </tr>
  <% 
do while not rsClass.eof 
%>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
    <td> <% 
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
	  response.write "<a href='Admin_Class_Article.asp?Action=Modify&ClassID=" & rsClass("ClassID") & "' title='" & rsClass("ReadMe") & "'>" & rsClass("ClassName") & "</a>"
	  if rsClass("Child")>0 then 
	  	response.write "��" & rsClass("Child") & "��" 
	  end if
	  
	  
	  'response.write "&nbsp;&nbsp;" & rsClass("ClassID") & "," & rsClass("PrevID") & "," & rsClass("NextID") & "," & rsClass("ParentID") & "," & rsClass("RootID")
	  %> </td>

    <td width="80" align="center"> <%
	if rsClass("LinkUrl")<>"" then
		response.write "<font color=red>�ⲿ</font>��"
	else
		response.write "<font color=green>ϵͳ</font>��"
	end if
	if rsClass("IsElite")=True then
		response.write "<font color=blue>�Ƽ�</font>"
	else
		response.write "��ͨ"
	end if
	%> </td>
    <td align="center"> <%
	select case rsClass("BrowsePurview")
	case 9999
		response.write "�ο�"
	case 999
		response.write "ע���û�"
	case 99
		response.write "�շ��û�"
	case 9
		response.write "VIP�û�"
	case 5
		response.write "����Ա"
	end select%> </td>
    <td align="center"><a href="Admin_Class_Article.asp?Action=Add&ParentID=<%=rsClass("ClassID")%>">�������Ŀ</a> 
      | <a href="Admin_Class_Article.asp?Action=Modify&ClassID=<%=rsClass("ClassID")%>">�޸�����</a> 
      | <a href="Admin_Class_Article.asp?Action=Move&ClassID=<%=rsClass("ClassID")%>">�ƶ���Ŀ</a> 
      | <a href="Admin_Class_Article.asp?Action=Clear&ClassID=<%=rsClass("ClassID")%>" onClick="return ConfirmDel3();">���</a> 
      | <a href="Admin_Class_Article.asp?Action=Del&ClassID=<%=rsClass("ClassID")%>" onClick="<%if rsClass("Child")>0 then%>return ConfirmDel1();<%else%>return ConfirmDel2();<%end if%>">ɾ��</a></td>
  </tr>
  <% 
	rsClass.movenext 
loop 
%>
</table> 
<script language="JavaScript" type="text/JavaScript">
function ConfirmDel1()
{
   alert("����Ŀ�»�������Ŀ��������ɾ����������Ŀ�����ɾ������Ŀ��");
   return false;
}

function ConfirmDel2()
{
   if(confirm("ɾ����Ŀ��ͬʱɾ������Ŀ�е��������£����Ҳ��ָܻ���ȷ��Ҫɾ������Ŀ��"))
     return true;
   else
     return false;
	 
}
function ConfirmDel3()
{
   if(confirm("�����Ŀ������Ŀ����������Ŀ�����������·������վ�У�ȷ��Ҫ��մ���Ŀ��"))
     return true;
   else
     return false;
	 
}
</script>
<%
end sub

sub AddClass()
%>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
  <form name="form1" method="post" action="Admin_Class_Article.asp" onSubmit="return check()">
	<tr>
	  <td colspan="3" align="center" class="title"><strong>�� �� �� �� �� Ŀ</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>������Ŀ��</strong><br>
        ����ָ��Ϊ�ⲿ��Ŀ </td>
      <td> <select name="ParentID">
          <%call Admin_ShowClass_Option(0,ParentID)%>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>��Ŀ���ƣ�</strong></td>
      <td><input name="ClassName" type="text" size="37" maxlength="20"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>��Ŀ˵����<br>
        </strong> ���������Ŀ������ʱ����ʾ�趨��˵�����֣���֧��HTML��</td>
      <td><textarea name="Readme" cols="37" rows="4" id="Readme"></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td><strong>�Ƿ����Ƽ���Ŀ��</strong><br>
        �Ƽ���Ŀ������ҳ������Ŀ�ĸ���Ŀ����ʾ�����б�</td>
      <td><input name="IsElite" type="radio" value="Yes" checked style="border: 0px;background-color: #eeeeee;">
        ��&nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="IsElite" value="No" style="border: 0px;background-color: #eeeeee;">
        �� </td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>��ĿͼƬ��ַ��</strong><br>
        ͼƬ����ʾ����Ŀǰ�档ע��ͼƬ��С��</td>
      <td><input name="ClassPicUrl" type="text" id="ClassPicUrl" size="37" maxlength="255">
        ��Ԥ�����ܣ�</td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>��Ŀ���ӵ�ַ��</strong><br>
        ����뽫��Ŀ���ӵ��ⲿ��ַ��������������URL��ַ�������뱣��Ϊ�ա�</td>
      <td><input name="LinkUrl" type="text" id="LinkUrl" size="37" maxlength="255"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>��Ŀ���Ȩ�ޣ�</strong><br>
        ֻ�о�����ӦȨ�޵��˲����������Ŀ�е����¡�</td>
      <td><select name="BrowsePurview" id="BrowsePurview">
          <option value="9999" <%if BrowsePurview=9999 then response.write " selected"%>>�ο�</option>
          <option value="999" <%if BrowsePurview=999 then response.write " selected"%>>ע���û�</option>
          <option value="99" <%if BrowsePurview=99 then response.write " selected"%>>�շ��û�</option>
          <option value="9" <%if BrowsePurview=9 then response.write " selected"%>>VIP�û�</option>
          <option value="5" <%if BrowsePurview=5 then response.write " selected"%>>����Ա</option>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd"> 
        <input name="Add" type="submit" value=" ��&nbsp;&nbsp;�� " style="cursor: hand;background-color: #cccccc;"> 
        &nbsp; <input name="Cancel" type="button" id="Cancel" value=" ȡ �� " onClick="window.location.href='Admin_Class_Article.asp'" style="cursor: hand;background-color: #cccccc;"> 
        <%'%if SkinCount=0 then response.write "<li><font color=red>���������Ŀ��ɫģ��</font></li>"
		'if SkinCount=0 then response.write "<li><font color=red>���������Ŀ��Ŀ���ģ��</font></li>" %></td>
    </tr>
</form>
</table>
<script language="JavaScript" type="text/JavaScript">
function check()
{
  if (document.form1.ClassName.value=="")
  {
    alert("��Ŀ���Ʋ���Ϊ�գ�");
	document.form1.ClassName.focus();
	return false;
  }
}
</script>
<%
end sub

sub Modify()
	dim ClassID,sql,rsClass,i
	ClassID=trim(request("ClassID"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		ClassID=CLng(ClassID)
	end if
	sql="select * From "&jieducm&"_ArticleClass where ClassID=" & ClassID
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ������Ŀ��</li>"
	else
%>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form name="form1" method="post" action="Admin_Class_Article.asp" onSubmit="return check()">
	<tr>
	  <td colspan="3" align="center" class="title"><strong>�� �� �� �� �� Ŀ</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>������Ŀ��</strong><br>
        �������ı�������Ŀ����<a href='Admin_Class_Article.asp?Action=Move&ClassID=<%=ClassID%>'>����ƶ���Ŀ</a></td>
      <td> <%
	if rsClass("ParentID")<=0 then
	  	response.write "�ޣ���Ϊһ����Ŀ��"
	else
    	dim rsParentClass,sqlParentClass
		sqlParentClass="Select * From "&jieducm&"_ArticleClass where ClassID in (" & rsClass("ParentPath") & ") order by Depth"
		set rsParentClass=server.CreateObject("adodb.recordset")
		rsParentClass.open sqlParentClass,conn,1,1
		do while not rsParentClass.eof
			for i=1 to rsParentClass("Depth")
				response.write "&nbsp;&nbsp;&nbsp;"
			next
			if rsParentClass("Depth")>0 then
				response.write "��"
			end if
			response.write "&nbsp;" & rsParentClass("ClassName") & "<br>"
			rsParentClass.movenext
		loop
		rsParentClass.close
		set rsParentClass=nothing
	end if
	%> </select></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>��Ŀ���ƣ�</strong></td>
      <td><input name="ClassName" type="text" value="<%=rsClass("ClassName")%>" size="37" maxlength="20"> 
        <input name="ClassID" type="hidden" id="ClassID" value="<%=rsClass("ClassID")%>"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>��Ŀ˵����<br>
        </strong> ���������Ŀ������ʱ����ʾ�趨��˵�����֣���֧��HTML��</td>
      <td><textarea name="Readme" cols="37" rows="4" id="Readme"><%=rsClass("ReadMe")%></textarea></td>
    </tr>
    <tr class="tdbg">
      <td><strong>�Ƿ����Ƽ���Ŀ��</strong><br>
        �Ƽ���Ŀ������ҳ������Ŀ�ĸ���Ŀ����ʾ�����б�</td>
      <td> <input name="IsElite" type="radio" value="Yes" <%if rsClass("IsElite")=True then response.write " checked"%> style="border: 0px;background-color: #eeeeee;">
        ��&nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="IsElite" value="No" <%if rsClass("IsElite")=False then response.write " checked"%> style="border: 0px;background-color: #eeeeee;">
        �� </td>
    </tr>
    <tr class="tdbg">
      <td><strong>�Ƿ��ڶ�����������ʾ��</strong><br>
        ֻѡ��ֻ��һ����Ŀ��Ч��</td>
      <td><input name="ShowOnTop" type="radio" value="Yes" <%if rsClass("ShowOnTop")=True then response.write " checked"%> style="border: 0px;background-color: #eeeeee;">
        ��&nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="ShowOnTop" value="No" <%if rsClass("ShowOnTop")=False then response.write " checked"%> style="border: 0px;background-color: #eeeeee;">
        �� </td>
    </tr>
    <tr class="tdbg">
      <td><strong>��Ŀ�ڵ���������ҳ����ʾ��ʽ��</strong><br>
        ��ѡ��ֻ��һ����Ŀ��Ч</td>
      <td><select name="Setting" id="Setting">
          <option value="0" <%if rsClass("Setting")=0 then response.write " selected"%>>ͼƬ����+��ͨ����</option>
          <option value="1" <%if rsClass("Setting")=1 then response.write " selected"%>>ֻ��ʾ��ͨ����</option>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>��ĿͼƬ��ַ��</strong><br>
        ͼƬ����ʾ����Ŀǰ�档ע��ͼƬ��С��</td>
      <td><input name="ClassPicUrl" type="text" id="ClassPicUrl" value="<%=rsClass("ClassPicUrl")%>" size="37" maxlength="255">
        ��Ԥ�����ܣ�</td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>��Ŀ�༭��</strong><br>
        ����༭����<span class="style1">��|��</span>�ָ����磺����|dilys|sws2000<br>
        ������ӡ������ܱࡱ���ϼ���Ĺ���Ա<br>
        ����ԱȨ�޲���Ȩ�޼̳��ƶ�</td>
      <td><input name="ClassMaster" type="text" id="ClassMaster" value="<%=rsClass("ClassMaster")%>" size="37" maxlength="100" disabled> 
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>��Ŀ���ӵ�ַ��</strong><br>
        ����뽫��Ŀ���ӵ��ⲿ��ַ��������������URL��ַ�������뱣��Ϊ�ա�</td>
      <td><input name="LinkUrl" type="text" id="LinkUrl" value="<%=rsClass("LinkUrl")%>" size="37" maxlength="255"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>��Ŀ���Ȩ�ޣ�</strong><br>
        ֻ�о�����ӦȨ�޵��˲����������Ŀ�е����¡�</td>
      <td><select name="BrowsePurview" id="BrowsePurview">
          <option value="9999" <%if rsClass("BrowsePurview")=9999 then response.write " selected"%>>�ο�</option>
          <option value="999" <%if rsClass("BrowsePurview")=999 then response.write " selected"%>>ע���û�</option>
          <option value="99" <%if rsClass("BrowsePurview")=99 then response.write " selected"%>>�շ��û�</option>
          <option value="9" <%if rsClass("BrowsePurview")=9 then response.write " selected"%>>VIP�û�</option>
          <option value="5" <%if rsClass("BrowsePurview")=5 then response.write " selected"%>>����Ա</option>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>��Ŀ��������Ȩ�ޣ�</strong><br>
        ֻ�о�����ӦȨ�޵��˲����ڴ���Ŀ�з������¡�</td>
      <td><select name="AddPurview" id="AddPurview">
          <option value="999" <%if rsClass("AddPurview")=999 then response.write " selected"%>>ע���û�</option>
          <option value="99" <%if rsClass("AddPurview")=99 then response.write " selected"%>>�շ��û�</option>
          <option value="9" <%if rsClass("AddPurview")=9 then response.write " selected"%>>VIP�û�</option>
          <option value="5" <%if rsClass("AddPurview")=5 then response.write " selected"%>>����Ա</option>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> 
        <input name="Submit" type="submit" value=" &nbsp;�����޸Ľ��&nbsp; " style="cursor: hand;background-color: #cccccc;"> 
        &nbsp; <input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='Admin_Class_Article.asp'" style="cursor: hand;background-color: #cccccc;"> 
      </td>
    </tr>
</form>
</table>
<script language="JavaScript" type="text/JavaScript">
function check()
{
  if (document.form1.ClassName.value=="")
  {
    alert("��Ŀ���Ʋ���Ϊ�գ�");
	document.form1.ClassName.focus();
	return false;
  }
}
</script>
<%
	end if
	rsClass.close
	set rsClass=nothing
end sub

sub MoveClass()
	dim ClassID,sql,rsClass,i
	dim SkinID,LayoutID,BrowsePurview,AddPurview
	ClassID=trim(request("ClassID"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		ClassID=CLng(ClassID)
	end if
	
	sql="select * From "&jieducm&"_ArticleClass where ClassID=" & ClassID
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ������Ŀ��</li>"
	else
%>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form name="form1" method="post" action="Admin_Class_Article.asp">
	<tr>
	  <td colspan="3" align="center" class="title"><strong>�� �� �� �� �� Ŀ</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="200"><strong>��Ŀ���ƣ�</strong></td>
      <td><%=rsClass("ClassName")%> <input name="ClassID" type="hidden" id="ClassID" value="<%=rsClass("ClassID")%>"></td>
    </tr>
    <tr class="tdbg">
      <td width="200"><strong>��ǰ������Ŀ��</strong></td>
      <td>
        <%
	if rsClass("ParentID")<=0 then
	  	response.write "�ޣ���Ϊһ����Ŀ��"
	else
    	dim rsParent,sqlParent
		sqlParent="Select * From "&jieducm&"_ArticleClass where ClassID in (" & rsClass("ParentPath") & ") order by Depth"
		set rsParent=server.CreateObject("adodb.recordset")
		rsParent.open sqlParent,conn,1,1
		do while not rsParent.eof
			for i=1 to rsParent("Depth")
				response.write "&nbsp;&nbsp;&nbsp;"
			next
			if rsParent("Depth")>0 then
				response.write "��"
			end if
			response.write "&nbsp;" & rsParent("ClassName") & "<br>"
			rsParent.movenext
		loop
		rsParent.close
		set rsParent=nothing
	end if
	%>
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="200"><strong>�ƶ�����</strong><br>
        ����ָ��Ϊ��ǰ��Ŀ����������Ŀ<br>
        ����ָ��Ϊ�ⲿ��Ŀ</td>
      <td><select name="ParentID" size="2" style="height:300px;width:500px;"><%call Admin_ShowClass_Option(0,rsClass("ParentID"))%></select></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveMove"> 
        <input name="Submit" type="submit" value=" &nbsp;�����ƶ����&nbsp; " style="cursor: hand;background-color: #cccccc;">
        &nbsp; 
        <input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='Admin_Class_Article.asp'" style="cursor: hand;background-color: #cccccc;">
	 </td>
   </tr>
</form>  
</table>
<%
	end if
	rsClass.close
	set rsClass=nothing
end sub

sub Order() 
	dim sqlClass,rsClass,i,iCount,j 
	sqlClass="select * From "&jieducm&"_ArticleClass where ParentID=0 order by RootID" 
	set rsClass=server.CreateObject("adodb.recordset") 
	rsClass.open sqlClass,conn,1,1 
	iCount=rsClass.recordcount 
%>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
	<tr>
	  <td colspan="4" align="center" class="title"><strong>һ �� �� Ŀ �� ��</strong></td> 
  </tr> 
  <% 
j=1 
do while not rsClass.eof 
%> 
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
      <td width="200">&nbsp;<%=rsClass("ClassName")%></td> 
<% 
	if j>1 then 
  		response.write "<form action='Admin_Class_Article.asp?Action=UpOrder' method='post'><td width='150'>" 
		response.write "<select name=MoveNum size=1><option value=0>�����ƶ�</option>" 
		for i=1 to j-1 
			response.write "<option value="&i&">"&i&"</option>" 
		next 
		response.write "</select>" 
		response.write "<input type=hidden name=ClassID value="&rsClass("ClassID")&">"
		response.write "<input type=hidden name=cRootID value="&rsClass("RootID")&">&nbsp;<input type=submit name=Submit value=��&nbsp;�� style='cursor: hand;background-color: #cccccc;'>" 
		response.write "</td></form>" 
	else 
		response.write "<td width='150'>&nbsp;</td>" 
	end if 
	if iCount>j then 
  		response.write "<form action='Admin_Class_Article.asp?Action=DownOrder' method='post'><td width='150'>" 
		response.write "<select name=MoveNum size=1><option value=0>�����ƶ�</option>" 
		for i=1 to iCount-j 
			response.write "<option value="&i&">"&i&"</option>" 
		next 
		response.write "</select>" 
		response.write "<input type=hidden name=ClassID value="&rsClass("ClassID")&">"
		response.write "<input type=hidden name=cRootID value="&rsClass("RootID")&">&nbsp;<input type=submit name=Submit value=��&nbsp;�� style='cursor: hand;background-color: #cccccc;'>" 
		response.write "</td></form>" 
	else 
		response.write "<td width='150'>&nbsp;</td>" 
	end if 
%> 
      <td>&nbsp;</td>
  </tr> 
  <% 
	j=j+1 
	rsClass.movenext 
loop 
%> 
</table> 
<% 
	rsClass.close 
	set rsClass=nothing 
end sub 

sub OrderN() 
	dim sqlClass,rsClass,i,iCount,trs,UpMoveNum,DownMoveNum 
	sqlClass="select * From "&jieducm&"_ArticleClass order by RootID,OrderID" 
	set rsClass=server.CreateObject("adodb.recordset") 
	rsClass.open sqlClass,conn,1,1 
%>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
	<tr>
	  <td colspan="4" align="center" class="title"><strong>N �� �� Ŀ �� ��</strong></td> 
  </tr> 
  <% 
do while not rsClass.eof 
%> 
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
      <td width="300"> 
	  <% 
	for i=1 to rsClass("Depth") 
	  	response.write "&nbsp;&nbsp;&nbsp;" 
	next 
	if rsClass("Child")>0 then 
		response.write "<img src='Images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>" 
	else 
	  	response.write "<img src='Images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>" 
	end if 
	if rsClass("ParentID")=0 then 
		response.write "<b>" 
	end if 
	response.write rsClass("ClassName") 
	if rsClass("Child")>0 then 
		response.write "(" & rsClass("Child") & ")" 
	end if 
	%></td> 
<% 
	if rsClass("ParentID")>0 then   '�������һ����Ŀ���������ͬ��ȵ���Ŀ��Ŀ���õ�����Ŀ����ͬ��ȵ���Ŀ������λ�ã�֮�ϻ���֮�µ���Ŀ���� 
		'��������������ӦΪFor i=1 to �ð�֮�ϵİ����� 
		set trs=nt2003.execute("select count(ClassID) From "&jieducm&"_ArticleClass where ParentID="&rsClass("ParentID")&" and OrderID<"&rsClass("OrderID")&"") 
		UpMoveNum=trs(0) 
		if isnull(UpMoveNum) then UpMoveNum=0 
		if UpMoveNum>0 then 
  			response.write "<form action='Admin_Class_Article.asp?Action=UpOrderN' method='post'><td width='150'>" 
			response.write "<select name=MoveNum size=1><option value=0>�����ƶ�</option>" 
			for i=1 to UpMoveNum 
				response.write "<option value="&i&">"&i&"</option>" 
			next 
			response.write "</select>" 
			response.write "<input type=hidden name=ClassID value="&rsClass("ClassID")&">&nbsp;<input type=submit name=Submit value=��&nbsp;�� style='cursor: hand;background-color: #cccccc;'>" 
			response.write "</td></form>" 
		else 
			response.write "<td width='150'>&nbsp;</td>" 
		end if 
		trs.close 
		'���ܽ���������ӦΪFor i=1 to �ð�֮�µİ����� 
		set trs=nt2003.execute("select count(ClassID) From "&jieducm&"_ArticleClass where ParentID="&rsClass("ParentID")&" and orderID>"&rsClass("orderID")&"") 
		DownMoveNum=trs(0) 
		if isnull(DownMoveNum) then DownMoveNum=0 
		if DownMoveNum>0 then 
  			response.write "<form action='Admin_Class_Article.asp?Action=DownOrderN' method='post'><td width='150'>" 
			response.write "<select name=MoveNum size=1><option value=0>�����ƶ�</option>" 
			for i=1 to DownMoveNum 
				response.write "<option value="&i&">"&i&"</option>" 
			next 
			response.write "</select>" 
			response.write "<input type=hidden name=ClassID value="&rsClass("ClassID")&">&nbsp;<input type=submit name=Submit value=��&nbsp;�� style='cursor: hand;background-color: #cccccc;'>" 
			response.write "</td></form>" 
		else 
			response.write "<td width='150'>&nbsp;</td>" 
		end if 
		trs.close 
	else 
		response.write "<td colspan=2>&nbsp;</td>" 
	end if 
%> 
      <td>&nbsp;</td>
  </tr> 
  <% 
	UpMoveNum=0 
	DownMoveNum=0 
	rsClass.movenext 
loop 
%> 
</table> 
<% 
	rsClass.close 
	set rsClass=nothing 
end sub 

sub Reset() 
%>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
  <form name="form1" method="post" action="Admin_Class_Article.asp?Action=SaveReset"> 
	<tr>
	  <td colspan="3" align="center" class="title"><strong>�� λ �� �� �� �� �� Ŀ</strong></td> 
  </tr> 
    <tr class="tdbg">  
    <td align="center">  
        <table width="80%" border="0" cellspacing="1" cellpadding="1"> 
          <tr class="tdbg">  
            <td height="150"><span class="style1"><strong>ע�⣺</strong></span><br> 
            &nbsp;&nbsp;&nbsp;&nbsp;���ѡ��λ������Ŀ����������Ŀ������Ϊһ����Ŀ����ʱ����Ҫ���¶Ը�����Ŀ���й����Ļ������á���Ҫ����ʹ�øù��ܣ����������˴�������ö��޷���ԭ��Ŀ֮��Ĺ�ϵ�������ʱ��ʹ�á�          
		    </td> 
          </tr> 
        </table> 
    <tr class="tdbg">  
    <td align="center">  
        <input type="submit" name="Submit" value="&nbsp;��λ������Ŀ&nbsp;" style="cursor: hand;background-color: #cccccc;"> &nbsp;&nbsp;&nbsp; 
		<input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='Admin_Class_Article.asp'" style="cursor: hand;background-color: #cccccc;">
      </td>
    </tr>
  </form>
</table>
<%
end sub

sub Unite()
%>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form name="myform" method="post" action="Admin_Class_Article.asp" onSubmit="return ConfirmUnite();">
	<tr>
	  <td colspan="3" align="center" class="title"><strong>�� �� �� Ŀ �� ��</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td align="center">
        &nbsp;&nbsp;����Ŀ 
        <select name="ClassID" id="ClassID">
        <%call Admin_ShowClass_Option(1,0)%>
        </select>
        �ϲ���
        <select name="TargetClassID" id="TargetClassID">
        <%call Admin_ShowClass_Option(1,0)%>
        </select>
	  </td>
	</tr>
  <tr class="tdbg"> 
    <td align="center">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="Action" type="hidden" id="Action" value="SaveUnite">
        <input type="submit" name="Submit" value=" &nbsp;�ϲ���Ŀ&nbsp; " style="cursor: hand;background-color: #cccccc;">
        &nbsp;&nbsp; 
        <input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='Admin_Class_Article.asp'" style="cursor: hand;background-color: #cccccc;">

	</td>
  </tr>
  <tr class="tdbg"> 
    <td height="60"><span class="style1"><strong>ע�����</strong></span><br>
      &nbsp;&nbsp;&nbsp;&nbsp;���в��������棬�����ز���������<br>
      &nbsp;&nbsp;&nbsp;&nbsp;������ͬһ����Ŀ�ڽ��в��������ܽ�һ����Ŀ�ϲ�����������Ŀ�С�Ŀ����Ŀ�в��ܺ�������Ŀ��<br>
    &nbsp;&nbsp;&nbsp;&nbsp;�ϲ�������ָ������Ŀ�����߰�����������Ŀ������ɾ�����������½�ת�Ƶ�Ŀ����Ŀ�С�</td>
  </tr>
 </form>
</table> 
<script language="JavaScript" type="text/JavaScript">
function ConfirmUnite()
{
  if (document.myform.ClassID.value==document.myform.TargetClassID.value)
  {
    alert("�벻Ҫ����ͬ��Ŀ�ڽ��в�����");
	document.myform.TargetClassID.focus();
	return false;
  }
  if (document.myform.TargetClassID.value=="")
  {
    alert("Ŀ����Ŀ����ָ��Ϊ��������Ŀ����Ŀ��");
	document.myform.TargetClassID.focus();
	return false;
  }
}
</script>
<% 
end sub 
%> 
<!--#include file="Admin_fooder.asp"-->
</body> 
</html> 
<% 
'�����Ŀ
sub SaveAdd()
	dim ClassID,ClassName,IsElite,ShowOnTop,Setting,Readme,ClassMaster,ClassPicUrl,LinkUrl,PrevOrderID
	dim sql,rs,trs
	dim RootID,ParentDepth,ParentPath,ParentStr,ParentName,MaxClassID,MaxRootID
	dim PrevID,NextID,Child

	ClassName=trim(request("ClassName"))
	ClassMaster=trim(request("ClassMaster"))
	IsElite=trim(request("IsElite"))
	ShowOnTop=trim(request("ShowOnTop"))
	'Setting=trim(request("Setting"))
	Setting=0
	Readme=trim(request("Readme"))
	ClassPicUrl=trim(request("ClassPicUrl"))
	LinkUrl=trim(request("LinkUrl"))
	'SkinID=Trim(request("SkinID"))
	SkinID=0
    'LayoutID=trim(request("LayoutID"))
	LayoutID=0
	BrowsePurview=trim(request("BrowsePurview"))
	AddPurview=trim(request("AddPurview"))
	if ClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��Ŀ���Ʋ���Ϊ�գ�</li>"
	end if
	if IsElite="Yes" then
		IsElite=True
	else
		IsElite=False
	end if
	if ShowOnTop="Yes" then
		ShowOnTop=True
	else
		ShowOnTop=False
	end if
	'if SkinID="" then
	'	FoundErr=True
	'	ErrMsg=ErrMsg & "<br><li>��ָ����Ŀ��ɫģ��</li>"
	'else
	'	SkinID=CLng(SkinID)
	'end if
	'if LayoutID="" then
	'	FoundErr=True
	'	ErrMsg=ErrMsg & "<br><li>��ָ���������ģ��</li>"
	'else
	'	LayoutID=CLng(LayoutID)
	'end if
	if ClassMaster<>"" then
		'call AddMaster(ClassMaster)
	end if
	if FoundErr=True then
		exit sub
	end if

	set rs = conn.execute("select Max(ClassID) From "&jieducm&"_ArticleClass")
	MaxClassID=rs(0)
	if isnull(MaxClassID) then
		MaxClassID=0
	end if
	rs.close
	ClassID=MaxClassID+1
	set rs=nt2003.execute("select max(rootid) From "&jieducm&"_ArticleClass")
	MaxRootID=rs(0)
	if isnull(MaxRootID) then
		MaxRootID=0
	end if
	rs.close
	RootID=MaxRootID+1
	
	if ParentID>0 then
		sql="select * From "&jieducm&"_ArticleClass where ClassID=" & ParentID & ""
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>������Ŀ�Ѿ���ɾ����</li>"
		else
			if rs("LinkUrl")<>"" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>����ָ���ⲿ��ĿΪ������Ŀ��</li>"
			end if
		end if
		if FoundErr=True then
			rs.close
			set rs=nothing
			exit sub
		else	
			RootID=rs("RootID")
			ParentName=rs("ClassName")
			ParentDepth=rs("Depth")
			ParentPath=rs("ParentPath")
			Child=rs("Child")
			ParentPath=ParentPath & "," & ParentID     '�õ�����Ŀ�ĸ�����Ŀ·��
			PrevOrderID=rs("OrderID")
			if Child>0 then		
				dim rsPrevOrderID
				'�õ��뱾��Ŀͬ�������һ����Ŀ��OrderID
				set rsPrevOrderID=nt2003.execute("select Max(OrderID) From "&jieducm&"_ArticleClass where ParentID=" & ParentID)
				PrevOrderID=rsPrevOrderID(0)
				set trs=nt2003.execute("select ClassID from "&jieducm&"_ArticleClass where ParentID=" & ParentID & " and OrderID=" & PrevOrderID)
				PrevID=trs(0)
				
				'�õ�ͬһ����Ŀ���ȱ���Ŀ�����������Ŀ�����OrderID�������ǰһ��ֵ����������ֵ��
				set rsPrevOrderID=nt2003.execute("select Max(OrderID) From "&jieducm&"_ArticleClass where ParentPath like '" & ParentPath & ",%'")
				if (not(rsPrevOrderID.bof and rsPrevOrderID.eof)) then
					if not IsNull(rsPrevOrderID(0))  then
				 		if rsPrevOrderID(0)>PrevOrderID then
							PrevOrderID=rsPrevOrderID(0)
						end if
					end if
				end if
			else
				PrevID=0
			end if

		end if
		rs.close
	else
		if MaxRootID>0 then
			set trs=nt2003.execute("select ClassID from "&jieducm&"_ArticleClass where RootID=" & MaxRootID & " and Depth=0")
			PrevID=trs(0)
			trs.close
		else
			PrevID=0
		end if
		PrevOrderID=0
		ParentPath="0"
	end if

	sql="Select * From "&jieducm&"_ArticleClass Where ParentID=" & ParentID & " AND ClassName='" & ClassName & "'"
	set rs=server.CreateObject("adodb.recordset")
	'Set Rs=Conn.Execute(Sql)
	rs.open sql,conn,1,1
	if not(rs.bof and rs.eof) then
		FoundErr=True
		if ParentID=0 then
			ErrMsg=ErrMsg & "<br><li>�Ѿ�����һ����Ŀ��" & ClassName & "</li>"
		else
			ErrMsg=ErrMsg & "<br><li>��" & ParentName & "�����Ѿ���������Ŀ��" & ClassName & "����</li>"
		end if
		rs.close
		set rs=nothing
		exit sub
	end if
	rs.close

	sql="Select top 1 * From "&jieducm&"_ArticleClass"
	rs.open sql,conn,1,3
    rs.addnew
	
   	rs("ClassName")=ClassName
	rs("IsElite")=IsElite
	rs("ShowOnTop")=ShowOnTop
	rs("Setting")=Clng(Setting)
	'rs("ClassMaster")=ClassMaster
	rs("RootID")=RootID
	rs("ParentID")=ParentID
	if ParentID>0 then
		rs("Depth")=ParentDepth+1
	else
		rs("Depth")=0
	end if
	rs("ParentPath")=ParentPath
	rs("OrderID")=PrevOrderID
	rs("Child")=0
	rs("Readme")=Readme
	rs("ClassPicUrl")=ClassPicUrl
	rs("LinkUrl")=LinkUrl
	rs("SkinID")=0
	rs("LayoutID")=0
	rs("BrowsePurview")=Cint(BrowsePurview)
    'rs("AddPurview")=Cint(AddPurview)
	AddPurview=9999
	rs("PrevID")=PrevID
	rs("NextID")=0
	rs.update
	rs.Close
    set rs=Nothing
	
	'�����뱾��Ŀͬһ����Ŀ����һ����Ŀ�ġ�NextID���ֶ�ֵ
	if PrevID>0 then
		conn.execute("update "&jieducm&"_ArticleClass set NextID=" & ClassID & " where ClassID=" & PrevID)
	end if
	
	if ParentID>0 then
		'�����丸�������Ŀ��
		conn.execute("update "&jieducm&"_ArticleClass set child=child+1 where ClassID="&ParentID)
		
		'���¸���Ŀ�����Լ����ڱ���Ҫ��ͬ�ڱ������µ���Ŀ�������
		conn.execute("update "&jieducm&"_ArticleClass set OrderID=OrderID+1 where rootid=" & rootid & " and OrderID>" & PrevOrderID)
		conn.execute("update "&jieducm&"_ArticleClass set OrderID=" & PrevOrderID & "+1 where ClassID=" & ClassID)
	end if
	
    'call CloseConn()
	Response.Redirect "Admin_Class_Article.asp"  
end sub

sub SaveModify()
	dim ClassName,Readme,IsElite,ShowOnTop,Setting,ClassMaster,ClassPicUrl,LinkUrl,SkinID,LayoutID,BrowsePurview,AddPurview
	dim trs,rs
	dim ClassID,sql,rsClass,i
	dim SkinCount,LayoutCount
	ClassID=trim(request("ClassID"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
	else
		ClassID=CLng(ClassID)
	end if
	ClassName=trim(request("ClassName"))
	IsElite=trim(request("IsElite"))
	ShowOnTop=trim(request("ShowOnTop"))
	Setting=trim(request("Setting"))
	ClassMaster=trim(request("ClassMaster"))
	Readme=trim(request("Readme"))
	ClassPicUrl=trim(request("ClassPicUrl"))
	LinkUrl=trim(request("LinkUrl"))
	SkinID=0
	LayoutID=0
	BrowsePurview=trim(request("BrowsePurview"))
	AddPurview=trim(request("AddPurview"))
	if ClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��Ŀ���Ʋ���Ϊ�գ�</li>"
	end if
	
	if FoundErr=True then
		exit sub
	end if
	
	sql="select * From "&jieducm&"_ArticleClass where ClassID=" & ClassID
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ������Ŀ��</li>"
		rsClass.close
		set rsClass=nothing
		exit sub
	end if
	if rsClass("Child")>0 and LinkUrl<>"" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>����Ŀ������Ŀ�����Բ�����Ϊ�ⲿ���ӵ�ַ��</li>"
	end if
	if IsElite="Yes" then
		IsElite=True
	else
		IsElite=False
	end if
	if ShowOnTop="Yes" then
		ShowOnTop=True
	else
		ShowOnTop=False
	end if

	if ClassMaster<>"" and ClassMaster<>rsClass("ClassMaster") then
		'call AddMaster(ClassMaster)
	end if
	
	if FoundErr=True then
		rsClass.close
		set rsClass=nothing
		exit sub
	end if
	
   	rsClass("ClassName")=ClassName
	rsClass("IsElite")=IsElite
	rsClass("ShowOnTop")=ShowOnTop
	rsClass("Setting")=Clng(Setting)
	'rsClass("ClassMaster")=ClassMaster
	rsClass("Readme")=Readme
	rsClass("ClassPicUrl")=ClassPicUrl
	rsClass("LinkUrl")=LinkUrl
	rsClass("SkinID")=0
	rsClass("LayoutID")=0
	rsClass("BrowsePurview")=Cint(BrowsePurview)
	rsClass("AddPurview")=Cint(AddPurview)
	rsClass.update
	rsClass.close
	set rsClass=nothing
	
	set rs=nothing
	set trs=nothing
	
    'call CloseConn()
	Response.Redirect "Admin_Class_Article.asp"  
end sub


sub DeleteClass()
	dim sql,rs,PrevID,NextID,ClassID
	ClassID=trim(Request("ClassID"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		ClassID=CLng(ClassID)
	end if
	
	sql="select ClassID,RootID,Depth,ParentID,Child,PrevID,NextID From "&jieducm&"_ArticleClass where ClassID="&ClassID
	set rs=server.CreateObject ("Adodb.recordset")
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��Ŀ�����ڣ������Ѿ���ɾ��</li>"
	else
		if rs("Child")>0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>����Ŀ��������Ŀ����ɾ��������Ŀ���ٽ���ɾ������Ŀ�Ĳ���</li>"
		end if
	end if
	if FoundErr=True then
		rs.close
		set rs=nothing
		exit sub
	end if
	PrevID=rs("PrevID")
	NextID=rs("NextID")
	if rs("Depth")>0 then
		conn.execute("update "&jieducm&"_ArticleClass set child=child-1 where ClassID=" & rs("ParentID"))
	end if
	rs.delete
	rs.update
	rs.close
	set rs=nothing
	'ɾ������Ŀ���������º�����
	conn.execute("delete from "&jieducm&"_Article where ClassID=" & ClassID)
	conn.execute("delete from "&jieducm&"_ArticleComment where ClassID=" & ClassID)
	
	'�޸���һ��Ŀ��NextID����һ��Ŀ��PrevID
	if PrevID>0 then
		conn.execute "update "&jieducm&"_ArticleClass set NextID=" & NextID & " where ClassID=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update "&jieducm&"_ArticleClass set PrevID=" & PrevID & " where ClassID=" & NextID
	end if
	'call CloseConn()
	response.redirect "Admin_Class_Article.asp"
		
end sub

sub ClearClass()
	dim strClassID,rs,trs,SuccessMsg,ClassID
	ClassID=trim(Request("ClassID"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		ClassID=CLng(ClassID)
	end if
	strClassID=cstr(ClassID)
	set rs=nt2003.execute("select ClassID,Child,ParentPath from "&jieducm&"_ArticleClass where ClassID=" & ClassID)
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��Ŀ�����ڣ������Ѿ���ɾ��</li>"
		exit sub
	end if
	if rs(1)>0 then
		set trs=nt2003.execute("select ClassID from "&jieducm&"_ArticleClass where ParentID=" & rs(0))
		do while not trs.eof
			strClassID=strClassID & "," & trs(0)
			trs.movenext
		loop
		trs.close
		set trs=nt2003.execute("select ClassID from "&jieducm&"_ArticleClass where ParentPath like '" & rs(2) & "," & rs(0) & ",%'")
		do while not trs.eof
			strClassID=strClassID & "," & trs(0)
			trs.movenext
		loop
		trs.close
		set trs=nothing
	end if
	rs.close
	set rs=nothing
	conn.execute("update "&jieducm&"_Article set Deleted=1 where ClassID in (" & strClassID & ")")
	conn.execute("delete from "&jieducm&"_Article where ClassID in (" & strClassID & ")")	
	SuccessMsg="����Ŀ����������Ŀ�������������Ѿ����Ƶ�����վ�У�"
	call WriteSuccessMsg(SuccessMsg)
end sub


sub SaveMove()
	dim ClassID,sql,rsClass,i
	dim rParentID
	dim trs,rs
	dim ParentID,RootID,Depth,Child,ParentPath,ParentName,iParentID,iParentPath,PrevOrderID,PrevID,NextID
	ClassID=trim(request("ClassID"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		ClassID=CLng(ClassID)
	end if
	
	sql="select * From "&jieducm&"_ArticleClass where ClassID=" & ClassID
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ������Ŀ��</li>"
		rsClass.close
		set rsClass=nothing
		exit sub
	end if

	rParentID=trim(request("ParentID"))
	if rParentID="" then
		rParentID=0
	else
		rParentID=CLng(rParentID)
	end if
	Response.Write(rParentID&"<br>")
	if rsClass("ParentID")<>rParentID then   '������������Ŀ����Ҫ��һϵ�м��
		if rParentID=rsClass("ClassID") then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>������Ŀ����Ϊ�Լ���</li>"
		end if
		'�ж���ָ������Ŀ�Ƿ�Ϊ�ⲿ��Ŀ����Ŀ��������Ŀ
		if rsClass("ParentID")=0 then
			if rParentID>0 then
				set trs=nt2003.execute("select rootid From "&jieducm&"_ArticleClass where LinkUrl='' and ClassID="&rParentID)
				if trs.bof and trs.eof then
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>����ָ���ⲿ��ĿΪ������Ŀ</li>"
				else
					if rsClass("rootid")=trs(0) then
						FoundErr=True
						ErrMsg=ErrMsg & "<br><li>����ָ������Ŀ��������Ŀ��Ϊ������Ŀ</li>"
					end if
				end if
				trs.close
				set trs=nothing
			end if
		else
			set trs=nt2003.execute("select ClassID From "&jieducm&"_ArticleClass where ParentPath like '"&rsClass("ParentPath")&"," & rsClass("ClassID") & "%' and ClassID="&rParentID)
			if not (trs.eof and trs.bof) then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>������ָ������Ŀ��������Ŀ��Ϊ������Ŀ</li>"
			end if
			trs.close
			set trs=nothing
		end if
		
	end if

	if FoundErr=True then
		rsClass.close
		set rsClass=nothing
		exit sub
	end if
	
	if rsClass("ParentID")=0 then
		ParentID=rsClass("ClassID")
		iParentID=0
	else
		ParentID=rsClass("ParentID")
		iParentID=rsClass("ParentID")
	end if
	Depth=rsClass("Depth")
	Child=rsClass("Child")
	RootID=rsClass("RootID")
	ParentPath=rsClass("ParentPath")
	PrevID=rsClass("PrevID")
	NextID=rsClass("NextID")
	rsClass.close
	set rsClass=nothing
	
	
  '���������������Ŀ
  '��Ҫ������ԭ��������Ŀ��Ϣ��������ȡ�����ID����Ŀ�������򡢼̳а���������
  '��Ҫ���µ�ǰ������Ŀ��Ϣ
  '�̳а���������Ҫ��д�������и���--ȡ������ǰ̨����ClassID in ParentPath�����
  dim mrs,MaxRootID
  set mrs=nt2003.execute("select max(rootid) From "&jieducm&"_ArticleClass")
  MaxRootID=mrs(0)
  set mrs=nothing
  if isnull(MaxRootID) then
	MaxRootID=0
  end if
  dim k,nParentPath,mParentPath
  dim ParentSql,ClassCount
  dim rsPrevOrderID
  if clng(parentid)<>rParentID and not (iParentID=0 and rParentID=0) then  '���������������Ŀ
  Response.Write(ParentID&"<br>")
  Response.Write(iParentID&"<br>")
  Response.Write(rParentID&"<br>")
  Response.End()
	'����ԭ��ͬһ����Ŀ����һ����Ŀ��NextID����һ����Ŀ��PrevID
	if PrevID>0 then
		conn.execute "update "&jieducm&"_ArticleClass set NextID=" & NextID & " where ClassID=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update "&jieducm&"_ArticleClass set PrevID=" & PrevID & " where ClassID=" & NextID
	end if
	
	if iParentID>0 and rParentID=0 then  	'���ԭ������һ������ĳ�һ������
		'�õ���һ��һ��������Ŀ
		sql="select ClassID,NextID from "&jieducm&"_ArticleClass where RootID=" & MaxRootID & " and Depth=0"
		set rs=server.CreateObject("Adodb.recordset")
		rs.open sql,conn,1,3
		PrevID=rs(0)      '�õ��µ�PrevID
		rs(1)=ClassID     '������һ��һ��������Ŀ��NextID��ֵ
		rs.update
		rs.close
		set rs=nothing
		
		MaxRootID=MaxRootID+1
		'���µ�ǰ��Ŀ����
		conn.execute("update "&jieducm&"_ArticleClass set depth=0,OrderID=0,rootid="&maxrootid&",parentid=0,ParentPath='0',PrevID=" & PrevID & ",NextID=0 where ClassID="&ClassID)
		'�����������Ŀ���������������Ŀ���ݡ�������Ŀ�������迼�ǣ�ֻ�����������Ŀ��Ⱥ�һ������ID(rootid)����
		if child>0 then
			i=0
			ParentPath=ParentPath & ","
			set rs=nt2003.execute("select * From "&jieducm&"_ArticleClass where ParentPath like '%"&ParentPath & ClassID&"%'")
			do while not rs.eof
				i=i+1				
				mParentPath=replace(rs("ParentPath"),ParentPath,"0,")
				conn.execute("update "&jieducm&"_ArticleClass set depth=depth-"&depth&",rootid="&maxrootid&",ParentPath='"&mParentPath&"' where ClassID="&rs("ClassID"))
				rs.movenext
			loop
			rs.close
			set rs=nothing
		end if
		
		'������ԭ��������Ŀ����Ŀ���������൱�ڼ�֦�����迼��
		conn.execute("update "&jieducm&"_ArticleClass set child=child-1 where ClassID="&iParentID)
		
	elseif iParentID>0 and rParentID>0 then    '����ǽ�һ������Ŀ�ƶ�����������Ŀ��
		'�õ���ǰ��Ŀ����������Ŀ��
		ParentPath=ParentPath & ","
		set rs=nt2003.execute("select count(*) From "&jieducm&"_ArticleClass where ParentPath like '%"&ParentPath & ClassID&"%'")
		ClassCount=rs(0)
		if isnull(ClassCount) then
			ClassCount=1
		end if
		rs.close
		set rs=nothing
		
		'���Ŀ����Ŀ�������Ϣ		
		set trs=nt2003.execute("select * From "&jieducm&"_ArticleClass where ClassID="&rParentID)
		if trs("Child")>0 then		
			'�õ��뱾��Ŀͬ�������һ����Ŀ��OrderID
			set rsPrevOrderID=nt2003.execute("select Max(OrderID) From "&jieducm&"_ArticleClass where ParentID=" & trs("ClassID"))
			PrevOrderID=rsPrevOrderID(0)
			'�õ��뱾��Ŀͬ�������һ����Ŀ��ClassID
			sql="select ClassID,NextID from "&jieducm&"_ArticleClass where ParentID=" & trs("ClassID") & " and OrderID=" & PrevOrderID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,3
			PrevID=rs(0)    '�õ��µ�PrevID
			rs(1)=ClassID     '������һ����Ŀ��NextID��ֵ
			rs.update
			rs.close
			set rs=nothing
			
			'�õ�ͬһ����Ŀ���ȱ���Ŀ�����������Ŀ�����OrderID�������ǰһ��ֵ����������ֵ��
			set rsPrevOrderID=nt2003.execute("select Max(OrderID) From "&jieducm&"_ArticleClass where ParentPath like '" & trs("ParentPath") & "," & trs("ClassID") & ",%'")
			if (not(rsPrevOrderID.bof and rsPrevOrderID.eof)) then
				if not IsNull(rsPrevOrderID(0))  then
			 		if rsPrevOrderID(0)>PrevOrderID then
						PrevOrderID=rsPrevOrderID(0)
					end if
				end if
			end if
		else
			PrevID=0
			PrevOrderID=trs("OrderID")
		end if
		
		'�ڻ���ƶ���������Ŀ�������������ָ����Ŀ֮�����Ŀ��������
		conn.execute("update "&jieducm&"_ArticleClass set OrderID=OrderID+" & ClassCount & "+1 where rootid=" & trs("rootid") & " and OrderID>" & PrevOrderID)
		
		'���µ�ǰ��Ŀ����
		conn.execute("update "&jieducm&"_ArticleClass set depth="&trs("depth")&"+1,OrderID="&PrevOrderID&"+1,rootid="&trs("rootid")&",ParentID="&rParentID&",ParentPath='" & trs("ParentPath") & "," & trs("ClassID") & "',PrevID=" & PrevID & ",NextID=0 where ClassID="&ClassID)
		
		'���������Ŀ���������Ŀ���ݣ����Ϊԭ���������ȼ��ϵ�ǰ������Ŀ�����
		set rs=nt2003.execute("select * From "&jieducm&"_ArticleClass where ParentPath like '%"&ParentPath&ClassID&"%' order by OrderID")
		i=1
		do while not rs.eof
			i=i+1
			iParentPath= trs("ParentPath")& "," & trs("ClassID") & "," & replace(rs("ParentPath"),ParentPath,"")
			conn.execute("update "&jieducm&"_ArticleClass set depth=depth-"&depth&"+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&iParentPath&"' where ClassID="&rs("ClassID"))
			rs.movenext
		loop
		rs.close
		set rs=nothing
		trs.close
		set trs=nothing
		
		'������ָ����ϼ���Ŀ������Ŀ��
		conn.execute("update "&jieducm&"_ArticleClass set child=child+1 where ClassID="&rParentID)
		
		'������ԭ���������Ŀ��			
		conn.execute("update "&jieducm&"_ArticleClass set child=child-1 where ClassID="&iParentID)
	else    '���ԭ����һ����Ŀ�ĳ�������Ŀ��������Ŀ
		'�õ��ƶ�����Ŀ����
		set rs=nt2003.execute("select count(*) From "&jieducm&"_ArticleClass where rootid="&rootid)
		ClassCount=rs(0)
		rs.close
		set rs=nothing
		
		'���Ŀ����Ŀ�������Ϣ		
		set trs=nt2003.execute("select * From "&jieducm&"_ArticleClass where ClassID="&rParentID)
		if trs("Child")>0 then		
			'�õ��뱾��Ŀͬ�������һ����Ŀ��OrderID
			set rsPrevOrderID=nt2003.execute("select Max(OrderID) From "&jieducm&"_ArticleClass where ParentID=" & trs("ClassID"))
			PrevOrderID=rsPrevOrderID(0)
			sql="select ClassID,NextID from "&jieducm&"_ArticleClass where ParentID=" & trs("ClassID") & " and OrderID=" & PrevOrderID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,3
			PrevID=rs(0)
			rs(1)=ClassID
			rs.update
			set rs=nothing
			
			'�õ�ͬһ����Ŀ���ȱ���Ŀ�����������Ŀ�����OrderID�������ǰһ��ֵ����������ֵ��
			set rsPrevOrderID=nt2003.execute("select Max(OrderID) From "&jieducm&"_ArticleClass where ParentPath like '" & trs("ParentPath") & "," & trs("ClassID") & ",%'")
			if (not(rsPrevOrderID.bof and rsPrevOrderID.eof)) then
				if not IsNull(rsPrevOrderID(0))  then
			 		if rsPrevOrderID(0)>PrevOrderID then
						PrevOrderID=rsPrevOrderID(0)
					end if
				end if
			end if
		else
			PrevID=0
			PrevOrderID=trs("OrderID")
		end if
	
		'�ڻ���ƶ���������Ŀ�������������ָ����Ŀ֮�����Ŀ��������
		conn.execute("update "&jieducm&"_ArticleClass set OrderID=OrderID+" & ClassCount &"+1 where rootid=" & trs("rootid") & " and OrderID>" & PrevOrderID)
		
		conn.execute("update "&jieducm&"_ArticleClass set PrevID=" & PrevID & ",NextID=0 where ClassID=" & ClassID)
		set rs=nt2003.execute("select * From "&jieducm&"_ArticleClass where rootid="&rootid&" order by OrderID")
		i=0
		do while not rs.eof
			i=i+1
			if rs("parentid")=0 then
				ParentPath=trs("ParentPath") & "," & trs("ClassID")
				conn.execute("update "&jieducm&"_ArticleClass set depth=depth+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&ParentPath&"',parentid="&rParentID&" where ClassID="&rs("ClassID"))
			else
				ParentPath=trs("ParentPath") & "," & trs("ClassID") & "," & replace(rs("ParentPath"),"0,","")
				conn.execute("update "&jieducm&"_ArticleClass set depth=depth+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&ParentPath&"' where ClassID="&rs("ClassID"))
			end if
			rs.movenext
		loop
		rs.close
		set rs=nothing
		trs.close
		set trs=nothing
		'������ָ����ϼ���Ŀ��Ŀ��		
		conn.execute("update "&jieducm&"_ArticleClass set child=child+1 where ClassID="&rParentID)

	end if
  end if
	
  'call CloseConn()
  Response.Redirect "Admin_Class_Article.asp"  
end sub

sub UpOrder()
	dim ClassID,sqlOrder,rsOrder,MoveNum,cRootID,tRootID,i,rs,PrevID,NextID
	ClassID=trim(request("ClassID"))
	cRootID=Trim(request("cRootID"))
	MoveNum=trim(request("MoveNum"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
	else
		ClassID=CLng(ClassID)
	end if
	if cRootID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>���������</li>"
	else
		cRootID=Cint(cRootID)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>���������</li>"
	else
		MoveNum=Cint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ���������֣�</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	'�õ�����Ŀ��PrevID,NextID
	set rs=nt2003.execute("select PrevID,NextID from "&jieducm&"_ArticleClass where ClassID=" & ClassID)
	PrevID=rs(0)
	NextID=rs(1)
	rs.close
	set rs=nothing
	'���޸���һ��Ŀ��NextID����һ��Ŀ��PrevID
	if PrevID>0 then
		conn.execute "update "&jieducm&"_ArticleClass set NextID=" & NextID & " where ClassID=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update "&jieducm&"_ArticleClass set PrevID=" & PrevID & " where ClassID=" & NextID
	end if

	dim mrs,MaxRootID
	set mrs=nt2003.execute("select max(rootid) From "&jieducm&"_ArticleClass")
	MaxRootID=mrs(0)+1
	'�Ƚ���ǰ��Ŀ������󣬰�������Ŀ
	conn.execute("update "&jieducm&"_ArticleClass set RootID=" & MaxRootID & " where RootID=" & cRootID)
	
	'Ȼ��λ�ڵ�ǰ��Ŀ���ϵ���Ŀ��RootID���μ�һ����ΧΪҪ����������
	sqlOrder="select * From "&jieducm&"_ArticleClass where ParentID=0 and RootID<" & cRootID & " order by RootID desc"
	set rsOrder=server.CreateObject("adodb.recordset")
	rsOrder.open sqlOrder,conn,1,3
	if rsOrder.bof and rsOrder.eof then
		exit sub        '�����ǰ��Ŀ�Ѿ��������棬�������ƶ�
	end if
	i=1
	do while not rsOrder.eof
		tRootID=rsOrder("RootID")       '�õ�Ҫ����λ�õ�RootID����������Ŀ
		conn.execute("update "&jieducm&"_ArticleClass set RootID=RootID+1 where RootID=" & tRootID)
		i=i+1
		if i>MoveNum then
			rsOrder("PrevID")=ClassID
			rsOrder.update
			conn.execute("update "&jieducm&"_ArticleClass set NextID=" & rsOrder("ClassID") & " where ClassID=" & ClassID)
			exit do
		end if
		rsOrder.movenext
	loop
	rsOrder.movenext
	if rsOrder.eof then
		conn.execute("update "&jieducm&"_ArticleClass set PrevID=0 where ClassID=" & ClassID)
	else
		rsOrder("NextID")=ClassID
		rsOrder.update
		conn.execute("update "&jieducm&"_ArticleClass set PrevID=" & rsOrder("ClassID") & " where ClassID=" & ClassID)
	end if	
	rsOrder.close
	set rsOrder=nothing
	
	'Ȼ���ٽ���ǰ��Ŀ������Ƶ���Ӧλ�ã���������Ŀ
	conn.execute("update "&jieducm&"_ArticleClass set RootID=" & tRootID & " where RootID=" & MaxRootID)
	'call CloseConn()
	response.Redirect "Admin_Class_Article.asp?Action=Order"
end sub

sub DownOrder()
	dim ClassID,sqlOrder,rsOrder,MoveNum,cRootID,tRootID,i,rs,PrevID,NextID
	ClassID=trim(request("ClassID"))
	cRootID=Trim(request("cRootID"))
	MoveNum=trim(request("MoveNum"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
	else
		ClassID=CLng(ClassID)
	end if
	if cRootID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>���������</li>"
	else
		cRootID=Cint(cRootID)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>���������</li>"
	else
		MoveNum=Cint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ���������֣�</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	'�õ�����Ŀ��PrevID,NextID
	set rs=nt2003.execute("select PrevID,NextID from "&jieducm&"_ArticleClass where ClassID=" & ClassID)
	PrevID=rs(0)
	NextID=rs(1)
	rs.close
	set rs=nothing
	'���޸���һ��Ŀ��NextID����һ��Ŀ��PrevID
	if PrevID>0 then
		conn.execute "update "&jieducm&"_ArticleClass set NextID=" & NextID & " where ClassID=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update "&jieducm&"_ArticleClass set PrevID=" & PrevID & " where ClassID=" & NextID
	end if

	dim mrs,MaxRootID
	set mrs=nt2003.execute("select max(rootid) From "&jieducm&"_ArticleClass")
	MaxRootID=mrs(0)+1
	'�Ƚ���ǰ��Ŀ������󣬰�������Ŀ
	conn.execute("update ArticleClass set RootID=" & MaxRootID & " where RootID=" & cRootID)
	
	'Ȼ��λ�ڵ�ǰ��Ŀ���µ���Ŀ��RootID���μ�һ����ΧΪҪ�½�������
	sqlOrder="select * From "&jieducm&"_ArticleClass where ParentID=0 and RootID>" & cRootID & " order by RootID"
	set rsOrder=server.CreateObject("adodb.recordset")
	rsOrder.open sqlOrder,conn,1,3
	if rsOrder.bof and rsOrder.eof then
		exit sub        '�����ǰ��Ŀ�Ѿ��������棬�������ƶ�
	end if
	i=1
	do while not rsOrder.eof
		tRootID=rsOrder("RootID")       '�õ�Ҫ����λ�õ�RootID����������Ŀ
		conn.execute("update "&jieducm&"_ArticleClass set RootID=RootID-1 where RootID=" & tRootID)
		i=i+1
		if i>MoveNum then
			rsOrder("NextID")=ClassID
			rsOrder.update
			conn.execute("update "&jieducm&"_ArticleClass set PrevID=" & rsOrder("ClassID") & " where ClassID=" & ClassID)
			exit do
		end if
		rsOrder.movenext
	loop
	rsOrder.movenext
	if rsOrder.eof then
		conn.execute("update "&jieducm&"_ArticleClass set NextID=0 where ClassID=" & ClassID)
	else
		rsOrder("PrevID")=ClassID
		rsOrder.update
		conn.execute("update "&jieducm&"_ArticleClass set NextID=" & rsOrder("ClassID") & " where ClassID=" & ClassID)
	end if	
	rsOrder.close
	set rsOrder=nothing
	
	'Ȼ���ٽ���ǰ��Ŀ������Ƶ���Ӧλ�ã���������Ŀ
	conn.execute("update "&jieducm&"_ArticleClass set RootID=" & tRootID & " where RootID=" & MaxRootID)
	'call CloseConn()
	response.Redirect "Admin_Class_Article.asp?Action=Order"
end sub

sub UpOrderN()
	dim sqlOrder,rsOrder,MoveNum,ClassID,i
	dim ParentID,OrderID,ParentPath,Child,PrevID,NextID
	ClassID=Trim(request("ClassID"))
	MoveNum=trim(request("MoveNum"))
	if ClassID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>���������</li>"
	else
		ClassID=CLng(ClassID)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>���������</li>"
	else
		MoveNum=Cint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ���������֣�</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	dim sql,rs,oldorders,ii,trs,tOrderID
	'Ҫ�ƶ�����Ŀ��Ϣ
	set rs=nt2003.execute("select ParentID,OrderID,ParentPath,child,PrevID,NextID From "&jieducm&"_ArticleClass where ClassID="&ClassID)
	ParentID=rs(0)
	OrderID=rs(1)
	ParentPath=rs(2) & "," & ClassID
	child=rs(3)
	PrevID=rs(4)
	NextID=rs(5)
	rs.close
	set rs=nothing
	if child>0 then
		set rs=nt2003.execute("select count(*) From "&jieducm&"_ArticleClass where ParentPath like '%"&ParentPath&"%'")
		oldorders=rs(0)
		rs.close
		set rs=nothing
	else
		oldorders=0
	end if
	'���޸���һ��Ŀ��NextID����һ��Ŀ��PrevID
	if PrevID>0 then
		conn.execute "update "&jieducm&"_ArticleClass set NextID=" & NextID & " where ClassID=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update "&jieducm&"_ArticleClass set PrevID=" & PrevID & " where ClassID=" & NextID
	end if
	
	'�͸���Ŀͬ������������֮�ϵ���Ŀ------���������򣬷�ΧΪҪ����������
	sql="select ClassID,OrderID,child,ParentPath,PrevID,NextID From "&jieducm&"_ArticleClass where ParentID="&ParentID&" and OrderID<"&OrderID&" order by OrderID desc"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
	i=1
	do while not rs.eof
		tOrderID=rs(1)
		conn.execute("update "&jieducm&"_ArticleClass set OrderID="&tOrderID+oldorders+i&" where ClassID="&rs(0))
		if rs(2)>0 then
			ii=i+1
			set trs=nt2003.execute("select ClassID,OrderID From "&jieducm&"_ArticleClass where ParentPath like '%"&rs(3)&","&rs(0)&"%' order by OrderID")
			if not (trs.eof and trs.bof) then
				do while not trs.eof
					conn.execute("update "&jieducm&"_ArticleClass set OrderID="&tOrderID+oldorders+ii&" where ClassID="&trs(0))
					ii=ii+1
					trs.movenext
				loop
			end if
			trs.close
			set trs=nothing
		end if
		i=i+1
		if i>MoveNum then
			rs(4)=ClassID
			rs.update
			conn.execute("update "&jieducm&"_ArticleClass set NextID=" & rs(0) & " where ClassID=" & ClassID)		
			exit do
		end if
		rs.movenext
	loop
	rs.movenext
	if rs.eof then
		conn.execute("update "&jieducm&"_ArticleClass set PrevID=0 where ClassID=" & ClassID)
	else
		rs(5)=ClassID
		rs.update
		conn.execute("update "&jieducm&"_ArticleClass set PrevID=" & rs(0) & " where ClassID=" & ClassID)
	end if	
	rs.close
	set rs=nothing
	
	'������Ҫ�������Ŀ�����
	conn.execute("update "&jieducm&"_ArticleClass set OrderID="&tOrderID&" where ClassID="&ClassID)
	'�����������Ŀ���������������Ŀ����
	if child>0 then
		i=1
		set rs=nt2003.execute("select ClassID From "&jieducm&"_ArticleClass where ParentPath like '%"&ParentPath&"%' order by OrderID")
		do while not rs.eof
			conn.execute("update "&jieducm&"_ArticleClass set OrderID="&tOrderID+i&" where ClassID="&rs(0))
			i=i+1
			rs.movenext
		loop
		rs.close
		set rs=nothing
	end if
	'call CloseConn()
	response.Redirect "Admin_Class_Article.asp?Action=OrderN"
end sub

sub DownOrderN()
	dim sqlOrder,rsOrder,MoveNum,ClassID,i
	dim ParentID,OrderID,ParentPath,Child,PrevID,NextID
	ClassID=Trim(request("ClassID"))
	MoveNum=trim(request("MoveNum"))
	if ClassID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>���������</li>"
		exit sub
	else
		ClassID=Cint(ClassID)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>���������</li>"
		exit sub
	else
		MoveNum=Cint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ�½������֣�</li>"
			exit sub
		end if
	end if

	dim sql,rs,oldorders,ii,trs,tOrderID
	'Ҫ�ƶ�����Ŀ��Ϣ
	set rs=nt2003.execute("select ParentID,OrderID,ParentPath,child,PrevID,NextID From "&jieducm&"_ArticleClass where ClassID="&ClassID)
	ParentID=rs(0)
	OrderID=rs(1)
	ParentPath=rs(2) & "," & ClassID
	child=rs(3)
	PrevID=rs(4)
	NextID=rs(5)
	rs.close
	set rs=nothing

	'���޸���һ��Ŀ��NextID����һ��Ŀ��PrevID
	if PrevID>0 then
		conn.execute "update "&jieducm&"_ArticleClass set NextID=" & NextID & " where ClassID=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update "&jieducm&"_ArticleClass set PrevID=" & PrevID & " where ClassID=" & NextID
	end if
	
	'�͸���Ŀͬ������������֮�µ���Ŀ------���������򣬷�ΧΪҪ�½�������
	sql="select ClassID,OrderID,child,ParentPath,PrevID,NextID From "&jieducm&"_ArticleClass where ParentID="&ParentID&" and OrderID>"&OrderID&" order by OrderID"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
	i=0      'ͬ����Ŀ
	ii=0     'ͬ����Ŀ������Ŀ
	do while not rs.eof
		conn.execute("update "&jieducm&"_ArticleClass set OrderID="&OrderID+ii&" where ClassID="&rs(0))
		if rs(2)>0 then
			set trs=nt2003.execute("select ClassID,OrderID From "&jieducm&"_ArticleClass where ParentPath like '%"&rs(3)&","&rs(0)&"%' order by OrderID")
			if not (trs.eof and trs.bof) then
				do while not trs.eof
					ii=ii+1
					conn.execute("update "&jieducm&"_ArticleClass set OrderID="&OrderID+ii&" where ClassID="&trs(0))
					trs.movenext
				loop
			end if
			trs.close
			set trs=nothing
		end if
		ii=ii+1
		i=i+1
		if i>=MoveNum then
			rs(5)=ClassID
			rs.update
			conn.execute("update "&jieducm&"_ArticleClass set PrevID=" & rs(0) & " where ClassID=" & ClassID)		
			exit do
		end if
		rs.movenext
	loop
	rs.movenext
	if rs.eof then
		conn.execute("update "&jieducm&"_ArticleClass set NextID=0 where ClassID=" & ClassID)
	else
		rs(4)=ClassID
		rs.update
		conn.execute("update "&jieducm&"_ArticleClass set NextID=" & rs(0) & " where ClassID=" & ClassID)
	end if	
	rs.close
	set rs=nothing
	
	'������Ҫ�������Ŀ�����
	conn.execute("update "&jieducm&"_ArticleClass set OrderID="&OrderID+ii&" where ClassID="&ClassID)
	'�����������Ŀ���������������Ŀ����
	if child>0 then
		i=1
		set rs=nt2003.execute("select ClassID From "&jieducm&"_ArticleClass where ParentPath like '%"&ParentPath&"%' order by OrderID")
		do while not rs.eof
			conn.execute("update "&jieducm&"_ArticleClass set OrderID="&OrderID+ii+i&" where ClassID="&rs(0))
			i=i+1
			rs.movenext
		loop
		rs.close
		set rs=nothing
	end if
	'call CloseConn()
	response.Redirect "Admin_Class_Article.asp?Action=OrderN"
end sub

sub SaveReset()
	dim i,sql,rs,SuccessMsg,iCount,PrevID,NextID
	sql="select ClassID From "&jieducm&"_ArticleClass order by RootID,OrderID"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
	iCount=rs.recordcount
	i=1
	PrevID=0
	do while not rs.eof
		rs.movenext
		if rs.eof then
			NextID=0
		else
			NextID=rs(0)
		end if
		rs.moveprevious
		conn.execute("update "&jieducm&"_ArticleClass set RootID=" & i & ",OrderID=0,ParentID=0,Child=0,ParentPath='0',Depth=0,PrevID=" & PrevID & ",NextID=" & NextID & " where ClassID=" & rs(0))
		PrevID=rs(0)
		i=i+1
		rs.movenext
	loop
	rs.close
	set rs=nothing	
	
	SuccessMsg="��λ�ɹ����뷵��<a href='Admin_Class_Article.asp'>��Ŀ������ҳ</a>����Ŀ�Ĺ������á�"
	call WriteSuccessMsg(SuccessMsg)
end sub

sub SaveUnite()
	dim ClassID,TargetClassID,ParentPath,iParentPath,Depth,iParentID,Child,PrevID,NextID
	dim rs,trs,i,SuccessMsg
	ClassID=trim(request("ClassID"))
	TargetClassID=trim(request("TargetClassID"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫ�ϲ�����Ŀ��</li>"
	else
		ClassID=CLng(ClassID)
	end if
	if TargetClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��Ŀ����Ŀ��</li>"
	else
		TargetClassID=CLng(TargetClassID)
	end if
	if ClassID=TargetClassID then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�벻Ҫ����ͬ��Ŀ�ڽ��в���</li>"
	end if
	if FoundErr=True then
		exit sub
	end if
	'�ж�Ŀ����Ŀ�Ƿ�������Ŀ������У��򱨴�
	set rs=nt2003.execute("select Child from "&jieducm&"_ArticleClass where ClassID=" & TargetClassID)
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>Ŀ����Ŀ�����ڣ������Ѿ���ɾ����</li>"
	else
		if rs(0)>0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>Ŀ����Ŀ�к�������Ŀ�����ܺϲ���</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	'�õ���ǰ��Ŀ��Ϣ
	set rs=nt2003.execute("select ClassID,ParentID,ParentPath,PrevID,NextID,Depth from "&jieducm&"_ArticleClass where ClassID="&ClassID)
	iParentID=rs(1)
	Depth=rs(5)
	if iParentID=0 then
		ParentPath=rs(0)
	else
		ParentPath=rs(2) & "," & rs(0)
	end if
	iParentPath=rs(0)
	PrevID=rs(3)
	NextID=rs(4)
	
	'�ж��Ƿ��Ǻϲ�����������Ŀ��
	set rs=nt2003.execute("select ClassID from "&jieducm&"_ArticleClass where ClassID="&TargetClassID&" and ParentPath like '"&ParentPath&"%'")
	if not (rs.eof and rs.bof) then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>���ܽ�һ����Ŀ�ϲ�������������Ŀ��</li>"
		exit sub
	end if
	
	'�õ���ǰ��Ŀ��������ĿID
	set rs=nt2003.execute("select ClassID from "&jieducm&"_ArticleClass where ParentPath like '"&ParentPath&"%'")
	i=0
	if not (rs.eof and rs.bof) then
		do while not rs.eof
			iParentPath=iParentPath & "," & rs(0)
			i=i+1
			rs.movenext
		loop
	end if
	if i>0 then
		ParentPath=iParentPath
	else
		ParentPath=ClassID
	end if
	
	'���޸���һ��Ŀ��NextID����һ��Ŀ��PrevID
	if PrevID>0 then
		conn.execute "update "&jieducm&"_ArticleClass set NextID=" & NextID & " where ClassID=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update "&jieducm&"_ArticleClass set PrevID=" & PrevID & " where ClassID=" & NextID
	end if
	
	'�������¼�����������Ŀ
	conn.execute("update "&jieducm&"_Article set ClassID="&TargetClassID&" where ClassID in ("&ParentPath&")")
	conn.execute("update "&jieducm&"_ArticleComment set ClassID="&TargetClassID&" where ClassID in ("&ParentPath&")")
	
	'ɾ�����ϲ���Ŀ����������Ŀ
	conn.execute("delete from "&jieducm&"_ArticleClass where ClassID in ("&ParentPath&")")
	
	'������ԭ��������Ŀ������Ŀ���������൱�ڼ�֦�����迼��
	if Depth>0 then
		conn.execute("update "&jieducm&"_ArticleClass set Child=Child-1 where ClassID="&iParentID)
	end if
	
	SuccessMsg="��Ŀ�ϲ��ɹ����Ѿ������ϲ���Ŀ������������Ŀ����������ת��Ŀ����Ŀ�С�<br><br>ͬʱɾ���˱��ϲ�����Ŀ��������Ŀ��"
	call WriteSuccessMsg(SuccessMsg)
	set rs=nothing
	set trs=nothing
end sub

%> 
