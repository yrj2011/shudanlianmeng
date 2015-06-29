<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：捷度传媒
'官方主页：http://www.jieducm.cn
'技术支持：robin 
'QQ;616591415
'电话：13598006137
'版权：版权属于捷度传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2008－2011 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
option explicit
response.buffer=true	
Const PurviewLevel=1    '操作权限
Const CheckChannelID=0    '所属频道，0为不检测所属频道
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->

<%
dim ChannelID,UserID,rs,sql
dim Action,FoundErr,ErrMsg
dim AdminPurview_Channel
Action=Trim(request("Action"))
ChannelID=trim(request("ChannelID"))
UserID=trim(request("UserID"))
if ChannelID="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>请指定频道ID</li>"
else
	ChannelID=Clng(ChannelID)
end if
if Action="Modify" then
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要修改的用户ID</li>"
	else
		UserID=Clng(UserID)
		sql="Select * from Admin where ID=" & UserID
		If Not IsObject(Conn) Then ConnectionDatabase 'shiyu
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.Open sql,conn,1,3
		if rs.Bof and rs.EOF then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>不存在此用户！</li>"
		else
			AdminName=rs("UserName")
			AdminPurview=rs("Purview")
			select case ChannelID
			case 2
				AdminPurview_Channel=rs("AdminPurview_Article")
			case 3
				AdminPurview_Channel=rs("AdminPurview_Soft")
			case 4
				AdminPurview_Channel=rs("AdminPurview_Photo")
			end select
		end if
		rs.close
		set rs=nothing
	end if
end if
if FoundErr=True then
	call WriteErrMsg()
	'call CloseConn()
	response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
<title>设置栏目权限</title>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="0" class="border">
<%
dim arrShowLine(10)
for i=0 to ubound(arrShowLine)
	arrShowLine(i)=False
next
dim sqlClass,rsClass,i,iDepth
select case ChannelID
case 2
	sqlClass="select * From "&jieducm&"_ArticleClass order by RootID,OrderID"
case 3
	sqlClass="select * From "&jieducm&"_ArticleClass order by RootID,OrderID"
case 4
	sqlClass="select * From "&jieducm&"_ArticleClass order by RootID,OrderID"
end select
set rsClass=server.CreateObject("adodb.recordset")
rsClass.open sqlClass,conn,1,1
%>
<form name="myform" method="post" action="">
  <tr align="center">
  <td colspan="4" class="title"><b>栏 目 预 览</b></td>
  </tr>
  <tr class="tdbg">
    <td width="25%"><strong>栏目名称</strong></td>
    <td width="25%" height="22"><strong>录入</strong></td>
    <td width="25%"><strong>审核</strong></td>
    <td width="25%" height="22"><strong>管理</strong></td>
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
    <td width="30" align="center"><%
	response.write "<input name='Purview_Add' type='checkbox' id='Purview_Add' value='" & rsClass("ClassID") & "' style='border: 0px;background-color: #eeeeee;'"
	if Action="Modify" and AdminPurview_Channel=3 then
		if CheckClassMaster(rsClass("ClassInputer"),AdminName)=True then response.write " checked"
	end if
	response.write ">"
	%>
	</td>
    <td width="30" align="center"><%
	response.write "<input name='Purview_Check' type='checkbox' id='Purview_Check' value='" & rsClass("ClassID") & "' style='border: 0px;background-color: #eeeeee;'"
	if Action="Modify" and AdminPurview_Channel=3 then
		if CheckClassMaster(rsClass("ClassChecker"),AdminName)=True then response.write " checked"
	end if
	response.write ">"
	%></td>
    <td width="30" align="center"><%
	response.write "<input name='Purview_Manage' type='checkbox' id='Purview_Manage' value='" & rsClass("ClassID") & "' style='border: 0px;background-color: #eeeeee;'"
	if Action="Modify" and AdminPurview_Channel=3 then
		if CheckClassMaster(rsClass("ClassMaster"),AdminName)=True then response.write " checked"
	end if
	response.write ">"
	%></td>
  </tr>
<% 
	rsClass.movenext 
loop 
%>
</form>
</table>
</body>
</html>
<%
rsClass.close
set rsClass=nothing
'call CloseConn()
%>
