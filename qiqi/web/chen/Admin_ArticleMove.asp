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
		ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ�ƶ������£�</li>"
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
		ErrMsg=ErrMsg & "<br><li>�Ҳ���Ҫ�Ƶ�������</li>"
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
<title>���ߴ�ý</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>
<body>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<form name="form1" method="post" action="Admin_ArticleMove.asp">
	  <tr>
	    <td colspan="3" align="center" class="title"><strong>�� �� �� ��</strong></td>
      </tr>
        <tr class="tdbg"> 
            <td width="100" align="right"><strong>������Ŀ��</strong></td>
            <td><%call Admin_ShowPath2(ParentPath,ClassName,Depth)%></td>
          </tr>
          <tr class="tdbg">
            <td width="100" align="right"><strong>����ר�⣺</strong></td>
            <td><%call Admin_ShowSpecial_Option(1,SpecialID)%></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>���±��⣺</strong></td>
            <td><%=Title%><input name="ArticleID" type="hidden" id="ArticleID" value="<%=ArticleID%>"></td>
          </tr>
          <tr class="tdbg"> 
            <td width="100" align="right"><strong>�ƶ�����</strong><br><br>
			   ����ָ��Ϊ��������Ŀ����Ŀ��ָ��Ϊ�ⲿ��Ŀ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            </td>
            <td><select name="TargetClassID" size="2"  style="height:300px;width:400px;">
                <%call Admin_ShowClass_Option(3,ClassID)%>
              </select> </td>
          </tr>
        <tr class="tdbg">
    <td align="center" colspan="2"><input name="strComeUrl" type="hidden" id="strComeUrl" value="<%=strComeUrl%>">
        <input name="Action" type="hidden" id="Action" value="Move">
        <input type="submit" name="Submit" value=" ȷ&nbsp;&nbsp;�� " style="cursor: hand;background-color: #cccccc;">  &nbsp;   &nbsp;   &nbsp; 
        <input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='Admin_ArticleManage.asp'" style="cursor: hand;background-color: #cccccc;"></td>
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
		ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ�ƶ������£�</li>"
	end if
	if TargetClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��Ŀ����Ŀ��</li>"
	else
		TargetClassID=CLng(TargetClassID)
		set tClass=nt2003.execute("select Child from "&jieducm&"_ArticleClass where ClassID=" & TargetClassID)
		if tClass.bof and tClass.eof then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ����Ŀ����Ŀ��</li>"
		else
			if tClass(0)>0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>���ܽ������Ƶ�������Ŀ����Ŀ�У�</li>"
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
		ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ�ƶ������£�</li>"
	end if
	if TargetClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��Ŀ����Ŀ��</li>"
	else
		TargetClassID=CLng(TargetClassID)
		set tClass=nt2003.execute("select Child from "&jieducm&"_ArticleClass where ClassID=" & TargetClassID)
		if tClass.bof and tClass.eof then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ����Ŀ����Ŀ��</li>"
		else
			if tClass(0)>0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>���ܽ������Ƶ�������Ŀ����Ŀ�У�</li>"
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
		ErrMsg=ErrMsg & "<br><li>��ָ��Ŀ��ר�⣡</li>"
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
