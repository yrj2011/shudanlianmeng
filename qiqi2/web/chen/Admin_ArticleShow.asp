<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
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
Const PurviewLevel=2
Const CheckChannelID=2
Const PurviewLevel_Article=3
%>
<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/admin_code_article.asp"-->
<SCRIPT language=javascript>
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ��ѡ�е������𣿱���������ѡ�е������Ƶ�����վ�С���Ҫʱ���ɴӻ���վ�лָ���"))
     return true;
   else
     return false;
}
</SCRIPT>
<%
dim ArticleID,sql,rs,FoundErr,ErrMsg,PurviewChecked,PurviewChecked2
dim ClassID,tClass,ClassName,RootID,ParentID,ParentPath,Depth,ClassMaster,ClassChecker
ArticleID=trim(request("ArticleID"))
FoundErr=False
PurviewChecked=False
PurviewChecked2=False

call main()
if FoundErr=True then
	WriteErrMsg()
end if
''call CloseConn()

sub main()
if ArticleId="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>��������</li>"
	exit sub
else
	ArticleID=Clng(ArticleID)
end if
sql="select * from "&jieducm&"_article where Deleted=0 and ArticleID=" & ArticleID & ""
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�Ҳ�������</li>"
else
	ClassID=rs("ClassID")
	set tClass=nt2003.execute("select ClassName,RootID,ParentID,Depth,ParentPath,ClassMaster,ClassChecker From "&jieducm&"_ArticleClass where ClassID=" & ClassID)
	if tClass.bof and tClass.eof then
		founderr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ������Ŀ</li>"
	else
		ClassName=tClass(0)
		RootID=tClass(1)
		ParentID=tClass(2)
		Depth=tClass(3)
		ParentPath=tClass(4)
		ClassMaster=tClass(5)
		ClassChecker=tClass(6)
	end if
	set tClass=nothing
end if
if FoundErr=True then
	rs.close
	set rs=nothing
	exit sub
end if
if AdminPurview=1 or AdminPurview_Article<=2 then
	PurviewChecked=True
else
	PurviewChecked=CheckClassMaster(ClassMaster,AdminName)
	if PurviewChecked=False and ParentID>0 then
		set tClass=nt2003.execute("select ClassMaster from "&jieducm&"_ArticleClass where ClassID in (" & ParentPath & ")")
		do while not tClass.eof
			PurviewChecked=CheckClassMaster(tClass(0),AdminName)
			if PurviewChecked=True then exit do
			tClass.movenext
		loop
	end if
	PurviewChecked2=CheckClassMaster(ClassChecker,AdminName)
	if PurviewChecked2=False and ParentID>0 then
		set tClass=nt2003.execute("select ClassMaster from "&jieducm&"_ArticleClass where ClassID in (" & ParentPath & ")")
		do while not tClass.eof
			PurviewChecked2=CheckClassMaster(tClass(0),AdminName)
			if PurviewChecked2=True then exit do
			tClass.movenext
		loop
	end if
end if
%>
<html>
<head>
<title><%=rs("Title")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" class="title" colspan="2" align="center"><b>�� �� Ԥ ��</b></td>
	</tr>
	<tr class="tdbg">
	<td>
<%
response.write "�����ڵ�λ�ã�&nbsp;<a href='Admin_ArticleManage.asp'>���¹���</a>&nbsp;&gt;&gt;&nbsp;"
if ParentID>0 then
	dim sqlPath,rsPath
	sqlPath="select ClassID,ClassName From "&jieducm&"_ArticleClass where ClassID in (" & ParentPath & ") order by Depth"
	set rsPath=server.createobject("adodb.recordset")
	rsPath.open sqlPath,conn,1,1
	do while not rsPath.eof
		response.Write "<a href='Admin_ArticleManage.asp?ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
		rsPath.movenext
	loop
	rsPath.close
	set rsPath=nothing
end if
response.write "<a href='Admin_ArticleManage.asp?ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
response.write "<a href='Admin_ArticleShow.asp?ArticleID=" & rs("ArticleID") & "'><font color='#ff6600'>" & rs("Title") & "</font></a>"
%>	
	</td>
    <td width="100" height="22" align="right">
<%
if rs("OnTop")=true then
	response.Write("<font color=blue>��</font>&nbsp;")
else
	response.write("&nbsp;&nbsp;&nbsp;")
end if
if rs("Hits")>=nt2003.site_setting(14) then
	response.write("<font color='#0066cc'>��</font>&nbsp;")
else
	response.write("&nbsp;&nbsp;&nbsp;")
end if
if rs("Elite")=true then
	response.write("<font color=green>��</font>")
else
	response.write("&nbsp;&nbsp;")
end if
%>
    </td>
  </tr>
  <tr align="center" class="tdbg"> 
    <td height="40" colspan="2" valign="bottom"><font size="5"><%=rs("Title")%></font></td>
  </tr>
  <% if  rs("Title1")<>"" then %>
<tr align="center" class="tdbg" valign="middle"> 
          <td height="29" colspan="2" valign="top"><font size="2"><%=rs("Title1")%></font></td>
         </tr>
<% end if %>

  <tr align="center" class="tdbg">
    <td colspan="2">
        <%
		dim Author,CopyFrom
		Author=rs("Author")
		CopyFrom=rs("CopyFrom")
		response.write "���ߣ�"
		if instr(Author,"|")>0 then
			response.write "<a href='mailto:" & right(Author,len(Author)-instr(Author,"|")) & "'>" & left(Author,instr(Author,"|")-1) & "</a>"
		else
			response.write Author
		end if
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;ת���ԣ�"
		if instr(CopyFrom,"|")>0 then
			response.write "<a href='" & right(CopyFrom,len(CopyFrom)-instr(CopyFrom,"|")) & "'>" & left(CopyFrom,instr(CopyFrom,"|")-1) & "</a>"
		else
			response.write CopyFrom
		end if
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;�������" & rs("Hits") & "&nbsp;&nbsp;&nbsp;&nbsp;����¼�룺<a href='Admin_ArticleManage.asp?Field=Editor&Keyword=" & rs("Editor") & "'>" & rs("Editor") & "</a>"
		%>
    </td>
  </tr>
  <tr class="tdbg">
    <td colspan="2"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
        <tr>
          <td height="200" valign="top"><p><%=rs("Content")%></p></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr class="tdbg"><td colspan="2">
      <li>��һƪ���£� 
        <%
	  dim rsPrev
	  sql="Select Top 1 ArticleID,Title From "&jieducm&"_Article Where Deleted=0 and ArticleID<" & rs("ArticleID") & " order by ArticleID desc"
	  Set rsPrev= Server.CreateObject("ADODB.Recordset")
	  rsPrev.open sql,conn,1,1
	  if rsPrev.Eof then
	  	response.write "û����"
	  else
	  	response.write "<a href='Admin_ArticleShow.asp?ArticleID="&rsPrev("ArticleID")& "'>"&rsPrev("Title") & "</a>"
	  end if
	  rsPrev.close
	  set rsPrev=nothing
	  %>
      </li>
      <br> <li> ��һƪ���£� 
        <%
	  dim rsNext
	  sql="Select Top 1 ArticleID,Title From "&jieducm&"_Article Where Deleted=0 and ArticleID>" & rs("ArticleID") & " order by ArticleID asc"
	  Set rsNext= Server.CreateObject("ADODB.Recordset")
	  rsNext.open sql,conn,1,1
	  if rsNext.Eof then
	  	response.write "û����"
	  else
	  	response.write "<a href='Admin_ArticleShow.asp?ArticleID="&rsNext("ArticleID")& "'>"&rsNext("Title") & "</a>"
	  end if
	  rsNext.close
	  set rsNext=nothing
	  %>
      </li></td>
  </tr>
  <tr align="right" class="tdbg"> 
    <td height="21" colspan="2">
<%
response.write "<strong>���ò�����</strong>"
if (rs("Editor")=AdminName and rs("Passed")=False) or PurviewChecked=True then
	response.write "<a href='Admin_ArticleModify.asp?ArticleID=" & rs("ArticleID") & "'>�޸�</a>&nbsp;&nbsp;"
    response.write "<a href='Admin_ArticleDel.asp?Action=Del&ArticleID=" & rs("ArticleID") & "' onclick='return ConfirmDel();'>ɾ��</a>&nbsp;&nbsp;" 
end if
if AdminPurview=1 or AdminPurview_Article<=2 then
	response.write "<a href='Admin_ArticleMove.asp?ArticleID=" & rs("ArticleID") & "'>�ƶ�</a>&nbsp;&nbsp;"
end if
if PurviewChecked2=True then
	if rs("Passed")=false then
		response.write "<a href='Admin_ArticleProperty.asp?Action=SetPassed&ArticleID=" & rs("ArticleID") & "'>ͨ�����</a>&nbsp;&nbsp;"
	Else
  		response.write "<a href='Admin_ArticleProperty.asp?Action=CancelPassed&ArticleID=" & rs("ArticleID") & "'>ȡ�����</a>&nbsp;&nbsp;"
  	end if
end if
if PurviewChecked=True then
  	if rs("OnTop")=false then
   		response.write "<a href='Admin_ArticleProperty.asp?Action=SetOnTop&ArticleID=" & rs("ArticleID") & "'>�̶�</a>&nbsp;&nbsp;"
   	else
		response.write "<a href='Admin_ArticleProperty.asp?Action=CancelOnTop&ArticleID=" & rs("ArticleID") & "'>���</a>&nbsp;&nbsp;" 
   	end if
  	if rs("Elite")=false then
   		response.write "<a href='Admin_ArticleProperty.asp?Action=SetElite&ArticleID=" & rs("ArticleID") & "'>��Ϊ�Ƽ�</a>"
   	else
		response.write "<a href='Admin_ArticleProperty.asp?Action=CancelElite&ArticleID=" & rs("ArticleID") & "'>ȡ���Ƽ�</a>"
    end if
end if
%></td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class="tdbg">
    <td class="title" colspan="6">
      <strong>������ۣ�</strong>
	</td>
<%
dim rsComment
sql="select * from "&jieducm&"_ArticleComment where ArticleID=" & ArticleID
Set rsComment= Server.CreateObject("ADODB.Recordset")
rsComment.open sql,conn,1,1
if rsComment.eof then
	response.write "<tr class='tdbg'><td>&nbsp;&nbsp;&nbsp;&nbsp;��ʱû���κ��˶Ա����·�������</td></tr>"
else
%>
        <tr align="center" class="title"> 
          <td width="30" height="22"><strong>ID</strong></td>
          <td width="516" height="22"><strong>����</strong></td>
          <td width="60" height="22"><strong>������</strong></td>
          <td width="120" height="22"><strong>������IP</strong></td>
          <td width="120" height="22"><strong>����ʱ��</strong></td>
          <td width="118" height="22"><strong>����</strong></td>
        </tr>
<%
	do while not rsComment.eof
%>
        <tr class="tdbg"> 
          <td width="30" align="center"><%= rsComment("CommentID") %></td>
          <td><% response.write "<a href=# title='" & rsComment("Content") & "'>" & left(rsComment("Content"),25) & "</a>" %></td>
          <td width="60" align="center"><%= rsComment("UserName") %></td>
          <td width="120" align="center"><%=rsComment("IP")%></td>
          <td width="120" align="center"><%= rsComment("WriteTime") %></td>
          <td width="118" align="center">
		  <%
		  if AdminPurview=1 or AdminPurview_Article=1 then
			  if rsComment("ReplyName")<>"" then
				  response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
			  else
				  response.write "<a href='Admin_ArticleComment.asp?Action=Reply&CommentID=" & rsComment("Commentid") & "'>�ظ�</a>&nbsp;&nbsp;"
			  end if
			  response.write "<a href='Admin_ArticleComment.asp?Action=Modify&CommentID=" & rsComment("Commentid") & "'>�޸�</a>&nbsp;&nbsp;"
			  response.write "<a href='Admin_ArticleComment.asp?Action=Del&CommentID=" & rsComment("CommentID") & "'>ɾ��</a>"
		  end if%>
		  </td>
        </tr>
        <%if rsComment("ReplyName")<>"" then%>
		<tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
          <td align="center">&nbsp;</td>
          <td colspan="4"><%response.write "����Ա��" & rsComment("ReplyName") & "���� " & rsComment("ReplyTime") & " �ظ���<br><div style='padding:0px 20px'>" & rsComment("ReplyContent") & "</div>"%></td>
          <td align="center"><a href="Admin_ArticleComment.asp?Action=Reply&CommentID=<%=rsComment("CommentID")%>">�޸Ļظ�����</a></td>
		</tr>
        <%
		end if
		rsComment.movenext
	loop
%>

<%
end if
rsComment.close
set rsComment=nothing
%>	</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" align="center" background=images/admin_bg_bottom.gif><font color=#cccccc>Copyright &copy; 2013 </font> <a href='http://www.778892.com' target="_blank"><font color=#ff6600><strong>���ߴ�ý</strong></font></a> <font color=#cccccc>. All Rights Reserved.</font></td>
  </tr>
</table>
</body>
</html>
<%
end sub
%>