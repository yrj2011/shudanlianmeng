<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：七  七 传 媒
'官方主页：http://www.778892.com
'技术支持：xsh
'QQ;859680966
'电话：15889679112
'版权：版权属于七七传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2013 七七传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
option explicit
response.buffer=true	
Const PurviewLevel=2
Const CheckChannelID=2
Const PurviewLevel_Article=2
%>
<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/admin_code_article.asp"-->
<%
dim ArticleID,ClassID,TargetClassID,SpecialID
dim Action,FoundErr,ErrMsg
dim SkinID,LayoutID,SkinCount,LayoutCount,ClassMaster,BrowsePurview,AddPurview
dim strComeUrl
if trim(request("strComeUrl"))="" then
	strComeUrl=Request.ServerVariables("HTTP_REFERER")
else
	strComeUrl=trim(request("strComeUrl"))
end if
ArticleID=trim(request("ArticleID"))
TargetClassID=trim(request("TargetClassID"))
SpecialID=trim(request("SpecialID"))
Action=Trim(Request("Action"))
if Action="MoveToClass" then
	call MoveToClass()
elseif Action="MoveToSpecial" then
	call MoveToSpecial()
elseif Action="Move" then
	call MoveArticle()
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
'call CloseConn()


sub main()
	if ArticleID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请选择要移动的文章！</li>"
		exit sub
	else
		ArticleID=CLng(ArticleID)
	end if
	dim sqlArticle,rsArticle
	dim Title,ClassName,ParentPath,Depth
	sqlArticle="select "&jieducm&"_Article.ClassID,"&jieducm&"_Article.Title,"&jieducm&"_ArticleClass.ClassName,"&jieducm&"_ArticleClass.ParentPath,"&jieducm&"_ArticleClass.Depth,"&jieducm&"_Article.SpecialID from "&jieducm&"_Article"
	sqlArticle=sqlArticle & " inner join "&jieducm&"_ArticleClass on "&jieducm&"_Article.ClassID="&jieducm&"_ArticleClass.ClassID where "&jieducm&"_Article.ArticleID=" & ArticleID
	set rsArticle=Server.CreateObject("ADODB.Recordset")
	rsArticle.open sqlArticle,conn,1,1
	if rsArticle.bof and rsArticle.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到要移到的文章</li>"
	else
		ClassID=rsArticle(0)
		Title=rsArticle(1)
		ClassName=rsArticle(2)
		ParentPath=rsArticle(3)
		Depth=rsArticle(4)
		SpecialID=rsArticle(5)		
%>
<html>
<head>
<title>七七传媒</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>
<body>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form name="form1" method="post" action="Admin_ArticleMove.asp">
	  <tr>
	    <td colspan="3" align="center" class="title"><strong>文 章 移 动</strong></td>
      </tr>
        <tr class="tdbg"> 
            <td width="100" align="right"><strong>所属栏目：</strong></td>
            <td><%call Admin_ShowPath2(ParentPath,ClassName,Depth)%></td>
          </tr>
          <tr class="tdbg">
            <td width="100" align="right"><strong>所属专题：</strong></td>
            <td><%call Admin_ShowSpecial_Option(1,SpecialID)%></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>文章标题：</strong></td>
            <td><%=Title%><input name="ArticleID" type="hidden" id="ArticleID" value="<%=ArticleID%>"></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>移动到：</strong><br><br>
			   不能指定为含有子栏目的栏目能指定为外部栏目&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            </td>
            <td><select name="TargetClassID" size="2"  style="height:300px;width:400px;">
                <%call Admin_ShowClass_Option(3,ClassID)%>
              </select> </td>
          </tr>
        <tr class="tdbg">
    <td align="center" colspan="2"><input name="strComeUrl" type="hidden" id="strComeUrl" value="<%=strComeUrl%>">
        <input name="Action" type="hidden" id="Action" value="Move">
        <input type="submit" name="Submit" value=" 确&nbsp;&nbsp;定 " style="cursor: hand;background-color: #cccccc;">  &nbsp;   &nbsp;   &nbsp; 
        <input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='Admin_ArticleManage.asp'" style="cursor: hand;background-color: #cccccc;"></td>
  </tr>
  </form>
</table>
<!--#include file="Admin_fooder.asp" -->
</table>
</body>
</html>
<%
	end if
	rsArticle.close
	set rsArticle=nothing
end sub

sub MoveArticle()
	dim tClass
	if ArticleID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请选择要移动的文章！</li>"
	end if
	if TargetClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定目标栏目！</li>"
	else
		TargetClassID=CLng(TargetClassID)
		set tClass=nt2003.execute("select Child from "&jieducm&"_ArticleClass where ClassID=" & TargetClassID)
		if tClass.bof and tClass.eof then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>找不到指定的目标栏目！</li>"
		else
			if tClass(0)>0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>不能将文章移到有子栏目的栏目中！</li>"
			end if
		end if			
	end if
	if FoundErr=True then exit sub
	if SpecialID="" then
		SpecialID=0
	else
		SpecialID=Clng(SpecialID)
	end if
	ArticleID=replace(ArticleID," ","")
	conn.execute("update "&jieducm&"_Article set ClassID=" & TargetClassID & ",SpecialID=" & SpecialID & ",Passed=0 where ArticleId in (" & ArticleID & ")")
	conn.execute("update "&jieducm&"_ArticleComment set ClassID=" & TargetClassID & " where ArticleId in (" & ArticleID & ")")
	'call CloseConn()
	response.Redirect strComeUrl
end sub

sub MoveToClass()
	dim tClass
	if ArticleID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请选择要移动的文章！</li>"
	end if
	if TargetClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定目标栏目！</li>"
	else
		TargetClassID=CLng(TargetClassID)
		set tClass=nt2003.execute("select Child from "&jieducm&"_ArticleClass where ClassID=" & TargetClassID)
		if tClass.bof and tClass.eof then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>找不到指定的目标栏目！</li>"
		else
			if tClass(0)>0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>不能将文章移到有子栏目的栏目中！</li>"
			end if
		end if			
	end if
	if FoundErr=True then exit sub
	ArticleID=replace(ArticleID," ","")
	conn.execute("update "&jieducm&"_Article set ClassID=" & TargetClassID & " where ArticleId in (" & ArticleID & ")")
	conn.execute("update "&jieducm&"_ArticleComment set ClassID=" & TargetClassID & " where ArticleId in (" & ArticleID & ")")
	'call CloseConn()
	response.Redirect strComeUrl
end sub

sub MoveToSpecial()
	if SpecialID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定目标专题！</li>"
		exit sub
	else
		SpecialID=CLng(SpecialID)
	end if
	ArticleID=replace(ArticleID," ","")
	conn.execute("update Article set SpecialID=" & SpecialID & " where ArticleID in (" & ArticleID & ")")
	'call CloseConn()
	response.Redirect strComeUrl
end sub

%>
