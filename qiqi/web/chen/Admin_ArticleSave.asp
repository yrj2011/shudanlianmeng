<%@language=vbscript codepage=936 %>
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
Const PurviewLevel_Article=3
%>
<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->

<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/admin_code_article.asp"-->
<%
nt2003.GetSite_Setting() 
dim Action,rs,sql,ErrMsg,FoundErr,ObjInstalled,fname,jieducm_name
dim ArticleID,ClassID,SpecialID,Title,Title1,Content,key,Author,CopyFrom,UpdateTime,Hits,Editor,mUserName
dim IncludePic,DefaultPicUrl,UploadFiles,Passed,OnTop,Hot,Elite,arrUploadFiles
dim TitleFontColor,TitleFontType,AuthorName,AuthorEmail,CopyFromName,CopyFromUrl,PaginationType,ReadLevel,ReadPoint,Stars,MaxCharPerPage
dim tClass,ClassName,Depth,ParentPath,Child,i
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
FoundErr=false
Action=trim(request("Action"))
ArticleID=Trim(Request.Form("ArticleID"))
ClassID=trim(request.form("ClassID"))
SpecialID=trim(request.Form("SpecialID"))
Title=Replace(trim(request.form("Title")),"&nbsp;"," ")
mUserName = Trim(Request.Form("mUserName"))
Title1=trim(request.form("Title1"))
TitleFontColor=trim(request.form("TitleFontColor"))
TitleFontType=trim(request.form("TitleFontType"))
Key=trim(request.form("Key"))
Content=trim(request.form("Content"))
Author=trim(request.form("Author"))
AuthorName=trim(request.form("AuthorName"))
AuthorEmail=trim(request.form("AuthorEmail"))
CopyFrom=trim(request.form("CopyFrom"))
CopyFromName=trim(request.form("CopyFromName"))
CopyFromUrl=trim(request.form("CopyFromUrl"))
UpdateTime=trim(request.form("UpdateTime"))
IncludePic=trim(request.form("IncludePic"))
DefaultPicUrl=trim(request.form("DefaultPicUrl"))
UploadFiles=trim(request.form("UploadFiles"))
Passed=trim(request.form("Passed"))
OnTop=trim(request.form("OnTop"))
Hot=trim(request.form("Hot"))
Elite=trim(request.form("Elite"))
Hits=trim(request.form("Hits"))
Editor=AdminName
PaginationType=trim(request.form("PaginationType"))
MaxCharPerPage=trim(request.form("MaxCharPerPage"))
ReadLevel=trim(request.form("ReadLevel"))
ReadPoint=trim(request.form("ReadPoint"))
Stars=trim(request.form("Stars"))
fname = now
fname=replace(fname,"/","")
fname = replace(fname,"-","")
fname = replace(fname," ","") 
fname = replace(fname,":","")
fname = replace(fname,"PM","")
fname = replace(fname,"AM","")
fname = replace(fname,"上午","")
fname = replace(fname,"下午","")
fname = fname&".htm"

if Action="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足</li>"
else
	call SaveArticle()
end if
if founderr=true then
	call WriteErrMsg()
else

	call SaveSuccess()
end if
'shiyu


sub SaveArticle()
	dim PurviewChecked
	if ClassID="" then
		founderr=true
		errmsg=errmsg & "<br><li>未指定图片所属栏目或者指定的栏目有下属子栏目</li>"
	else
		ClassID=CLng(ClassID)
		if ClassID<=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>指定了非法的栏目（外部栏目或不存在的栏目）</li>"
		else
			set tClass=nt2003.execute("select ClassName,Depth,ParentPath,Child,LinkUrl,ParentID,ClassInputer From "&jieducm&"_ArticleClass where ClassID=" & ClassID)
			if tClass.bof and tClass.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>找不到指定的栏目</li>"
			else
				ClassName=tClass(0)
				Depth=tClass(1)
				ParentPath=tClass(2)
				Child=tClass(3)
			'if nt2003.site_setting(43)<>1 then  '判断是否含子栏目可发表 一枝梅+
				if Child>0 then
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>指定的栏目有下属子栏目</li>"
				end if
			'end if   '结束
				if tClass(4)<>"" then
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>不能指定外部栏目</li>"
				end if
				if AdminPurview=2 and AdminPurview_Article=3 then
					PurviewChecked=CheckClassMaster(tClass(6),AdminName)
					if PurviewChecked=False and tClass(5)>0 then
						set tClass=nt2003.execute("select ClassInputer from "&jieducm&"_ArticleClass where ClassID in (" & ParentPath & ")")
						do while not tClass.eof
							PurviewChecked=CheckClassMaster(tClass(0),AdminName)
							if PurviewChecked=True then exit do
							tClass.movenext
						loop
					end if
					if PurviewChecked=False then
						FoundErr=True
						ErrMsg=ErrMsg & "<br><li>对不起，你没有在此栏目发表文章的权限！</li>"
					end if
				end if
			end if
		end if
	end if
	if Title="" then
		founderr=true
		errmsg=ErrMsg & "<br><li>文章标题不能为空</li>"
	end if
	if Key="" then
		founderr=true
		errmsg=errmsg & "<br><li>请输入文章关键字</li>"
	end if
	if Content="" then
		founderr=true
		errmsg=errmsg & "<br><li>文章内容不能为空</li>"
	end if
	if PaginationType="" then
		PaginationType=0
	else
		PaginationType=Cint(PaginationType)
	end if
	if MaxCharPerPage="" then
		MaxCharPerPage=0
	else
		MaxCharPerPage=CLng(MaxCharPerPage)
	end if
	if PaginationType=1 and MaxCharPerPage=0 then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定自动分页时的每页大约字符数，必须大于0</li>"
	end if
	
	if FoundErr=True then
		exit sub
	end if

	if SpecialID="" then
		SpecialID=0
	else
		SpecialID=CLng(SpecialID)
	end if
	Title=dvhtmlencode(Title)
	if TitleFontType="" then
		TitleFontType=0
	end if
	
	dim strSiteUrl
	strSiteUrl=request.ServerVariables("HTTP_REFERER")
	strSiteUrl=lcase(left(strSiteUrl,instrrev(strSiteUrl,"/")))
	Content=ubbcode(replace(Content,strSiteUrl,""))
'远程图片保存
if nt2003.site_setting(8)=1 then   '远程图片保存
	Content=ReplaceRemoteUrl(Content)
end if 
	if Author<>"" then
		Author=dvhtmlencode(Author)
	else
		if AuthorName="" and AuthorEmail="" then
			Author="佚名"
		else
			if AuthorName<>"" then
				Author=AuthorName
				if AuthorEmail<>"" then
			 		Author=Author & "|" & AuthorEmail
				end if
			end if
		end if
	end if
	if CopyFrom<>"" then
		CopyFrom=dvhtmlencode(CopyFrom)
	else
		if CopyFromName="" and CopyFromUrl="" then
			CopyFrom="本站原创"
		else
			if CopyFromName<>"" then
				CopyFrom=CopyFromName
				if CopyFromUrl<>"" then
					CopyFrom=CopyFrom & "|" & CopyFromUrl
				end if
			end if
		end if			
	end if
	if UpdateTime<>"" and IsDate(UpdateTime)=true then
		UpdateTime=CDate(UpdateTime)
	else
		UpdateTime=now()
	end if
	if Hits<>"" then
		Hits=CLng(Hits)
	else
		Hits=0
	end if
	if ReadLevel="" then
		ReadLevel=9999
	else
		ReadLevel=CInt(ReadLevel)
	end if
	if ReadPoint="" then
		ReadPoint=0
	else
		ReadPoint=Cint(ReadPoint)
	end if
	if Stars="" then
		Stars=0
	else
		Stars=CInt(Stars)
	end if
	
	set rs=server.createobject("adodb.recordset")
	if Action="Add1" or Action="Add2" then
		sql="select * from "&jieducm&"_article" 
		rs.open sql,conn,1,3
		rs.addnew
		call SaveData()
		if Editor <> "" Then rs("Editor")=Editor
		rs.update
		ArticleID=rs("ArticleID")
		'Response.Write(ArticleID)  'DEBUG
		'Response.End()             'BEBUG
		rs.close
	elseif Action="Modify" then
  		if ArticleID="" then
			founderr=true
			errmsg=errmsg & "<br><li>不能确定ArticleID的值</li>"
		else
			ArticleID=Clng(ArticleID)
			sql="select * from "&jieducm&"_article where articleid=" & ArticleID
			rs.open sql,conn,1,3
			if rs.bof and rs.eof then
				founderr=true
				errmsg=errmsg & "<br><li>找不到此文章，可能已经被其他人删除。</li>"
 			else
		
	if ClassID <> "" Then rs("ClassID")=ClassID
	if SpecialID <> "" Then rs("SpecialID")=SpecialID
	if Title <> "" Then rs("Title")=Title
	if mUserName <> "" Then rs("mUserName") = mUserName
	if Title1 <> "" Then rs("Title1")=Title1
	if TitleFontColor <> "" Then rs("TitleFontColor")=TitleFontColor
	if TitleFontType <> "" Then rs("TitleFontType")=TitleFontType
	if Content <> "" Then rs("Content")=Content
	rs("Key1")=request("key")
	rs("Receive")=0
	
	if Author <> "" Then rs("Author")=Author
	if CopyFrom <> "" Then rs("CopyFrom")=CopyFrom
	rs("Deleted")=0
	if IncludePic="yes" then
		rs("IncludePic")=True
	else
		rs("IncludePic")=False
	end if
	if Passed="yes" then
		rs("Passed")=1 'True Re 1
	else
		if nt2003.site_setting(5)=0 then
			rs("Passed")=1
		else
			rs("Passed")=0
		end if
	end if



	if Stars <> "" Then rs("Stars")=Stars
	if PaginationType <> "" Then rs("PaginationType")=PaginationType
	if MaxCharPerPage <> "" Then rs("MaxCharPerPage")=MaxCharPerPage
	if ReadLevel <> "" Then rs("ReadLevel")=ReadLevel
	if ReadPoint <> "" Then rs("ReadPoint")=ReadPoint
	rs("NewsPath") = fname

	if DefaultPicUrl <> "" Then rs("DefaultPicUrl")=DefaultPicUrl
	if UploadFiles <> "" Then rs("UploadFiles")=UploadFiles
				rs.update
				rs.close
			end if
		end if
	else
		FoundErr=True
		ErrMsg="<br><li>参数错误!</li>"
	end if
	set rs=nothing
end sub

sub SaveData()
	if ClassID <> "" Then rs("ClassID")=ClassID
	if SpecialID <> "" Then rs("SpecialID")=SpecialID
	if Title <> "" Then rs("Title")=Title
	if mUserName <> "" Then rs("mUserName") = mUserName
	if Title1 <> "" Then rs("Title1")=Title1
	if TitleFontColor <> "" Then rs("TitleFontColor")=TitleFontColor
	if TitleFontType <> "" Then rs("TitleFontType")=TitleFontType
	if Content <> "" Then rs("Content")=Content
	rs("Key1")=request("key")
	rs("Receive")=0
	rs("hits")=1
	if Author <> "" Then rs("Author")=Author
	if CopyFrom <> "" Then rs("CopyFrom")=CopyFrom
	rs("Deleted")=0
	if IncludePic="yes" then
		rs("IncludePic")=True
	else
		rs("IncludePic")=False
	end if
	if Passed="yes" then
		rs("Passed")=1 'True Re 1
	else
		if nt2003.site_setting(5)=0 then
			rs("Passed")=1
		else
			rs("Passed")=0
		end if
	end if

	if OnTop="yes" then
		rs("OnTop")=True
	else
		rs("OnTop")=False
	end if
	if Hot="yes" then
		rs("Hot")=True
	else
		rs("Hot")=False
	end if
	if Elite="yes" then
		rs("Elite")=True
	else
		rs("Elite")=False
	end if
	if Stars <> "" Then rs("Stars")=Stars
	if UpdateTime <> "" Then rs("UpdateTime")=UpdateTime
	if PaginationType <> "" Then rs("PaginationType")=PaginationType
	if MaxCharPerPage <> "" Then rs("MaxCharPerPage")=MaxCharPerPage
	if ReadLevel <> "" Then rs("ReadLevel")=ReadLevel
	if ReadPoint <> "" Then rs("ReadPoint")=ReadPoint
	rs("NewsPath") = fname

	'***************************************
	'删除无用的上传文件
	
	'结束
	'***************************************
	if DefaultPicUrl <> "" Then rs("DefaultPicUrl")=DefaultPicUrl
	if UploadFiles <> "" Then rs("UploadFiles")=UploadFiles
end sub
	
sub SaveSuccess()
dim orderid
set orderid=conn.execute("select top 1 ArticleID,Orderid from "&jieducm&"_Article where Title='"&title&"' order by ArticleID desc")
conn.execute("update "&jieducm&"_Article set orderid=articleid*100 where articleID="&orderid("articleID"))
%>
<html>
<head>
<title>七 七 传 媒</title>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table class="border" width="100%" border="0" cellpadding="0" cellspacing="1">
  <tr> 
    <td height="22" align="center" class="title" colspan="2"><b> 
      <%if action="Add1" or Action="Add2" then%>
      添加 
      <%else%>
      修改 
      <%end if%>
      文章成功</b></td>
  </tr>
<%if Passed<>"yes" then%>
  <tr class="tdbg"> 
    <td height="60" colspan="2">&nbsp;</td>
  </tr>
<%end if%>
        <tr> 
          <td width="100" align="right" class="tdbg"><strong>所属栏目：</strong></td>
          <td width="894" class="tdbg"><%call Admin_ShowPath2(ParentPath,ClassName,Depth)%></td>
        </tr>
        <tr> 
          <td width="100" align="right" class="tdbg"><strong>文章标题：</strong></td>
          <td class="tdbg"><%= Title %></td>
        </tr>
        <tr> 
          <td width="100" align="right" class="tdbg"><strong>副标题：</strong></td>
          <td class="tdbg"><%= Title1 %></td>
        </tr>
        <tr> 
          <td width="100" align="right" class="tdbg"><strong>作&nbsp;&nbsp;&nbsp;&nbsp;者：</strong></td>
          <td class="tdbg"><%= Author %></td>
        </tr>
        <tr> 
          <td width="100" align="right" class="tdbg"><strong>转 贴 自：</strong></td>
          <td class="tdbg"><%= CopyFrom %></td>
        </tr>
        <tr> 
          <td width="100" align="right" class="tdbg"><strong>关 键 字：</strong></td>
          <td class="tdbg"><%= mid(Key,2,len(Key)-2) %></td>
        </tr>
     </td>
  </tr>
  <tr> 
    <td height="30" align="center" colspan="2" class="tdbg">【<a href="Admin_ArticleModify.asp?ArticleID=<%=ArticleID%>">修改本文</a>】&nbsp;【<a href="<%if Action="Add1" then%>Admin_ArticleAdd1.asp<%else%>Admin_ArticleAdd1.asp<%end if%>">继续添加文章</a>】&nbsp;【<a href="Admin_ArticleManage.asp">文章管理</a>】&nbsp;【<a href="Admin_ArticleShow.asp?ArticleID=<%=ArticleID%>">预览文章内容</a>】
	</td>
  </tr>
</table>
<!--#include file="Admin_fooder.asp" -->
</table>
</body>
</html>
<%
	session("ClassID_Article")=ClassID
	session("SpecialID_Article")=SpecialID
	session("Key")=trim(request("Key"))
	session("Author")=Author
	session("AuthorName")=AuthorName
	session("AuthorEmail")=AuthorEmail
	session("CopyFrom")=CopyFrom
	session("CopyFromName")=CopyFromName
	session("CopyFromUrl")=CopyFromUrl
	session("PaginationType")=PaginationType
	session("ReadLevel")=ReadLevel
	session("ReadPoint")=ReadPoint
end sub


'==================================================
'过程名：ReplaceRemoteUrl
'作  用：替换字符串中的远程文件为本地文件并保存远程文件
'参  数：strContent ------ 要替换的字符串
'==================================================
function ReplaceRemoteUrl(strContent)
	if IsObjInstalled("Microsoft.XMLHTTP")=False or nt2003.site_setting(8)<>1 then
		ReplaceRemoteUrl=strContent
		exit function
	end if
			
	dim re,RemoteFile,RemoteFileurl,SaveFilePath,SaveFileName,SaveFileType,arrSaveFileName,ranNum
	SaveFilePath = "../admin/UploadFiles"			'文件保存的本地路径
	if right(SaveFilePath,1)<>"/" then SaveFilePath=SaveFilePath&"/"
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}((\w)+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(gif|jpg|png|bmp)))"
	Set RemoteFile = re.Execute(strContent)
	For Each RemoteFileurl in RemoteFile
		arrSaveFileName = split(RemoteFileurl,".")
		SaveFileType=arrSaveFileName(ubound(arrSaveFileName))
		ranNum=int(900*rnd)+100
		jieducm_name="jieducm_"
		SaveFileName = SaveFilePath&jieducm_name&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&SaveFileType	
		call SaveRemoteFile(SaveFileName,RemoteFileurl)
		strContent=Replace(strContent,RemoteFileurl,SaveFileName)
		if UploadFiles="" then
			UploadFiles=SaveFileName
		else
			UploadFiles=UploadFiles & "|" & SaveFileName
		end if
	Next
	ReplaceRemoteUrl=strContent
end function



'==================================================
'过程名：SaveRemoteFile
'作  用：保存远程的文件到本地
'参  数：LocalFileName ------ 本地文件名
'		 RemoteFileUrl ------ 远程文件URL
'==================================================
sub SaveRemoteFile(LocalFileName,RemoteFileUrl)
	dim Ads,Retrieval,GetRemoteData
	Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
	With Retrieval
		.Open "Get", RemoteFileUrl, False, "", ""
		.Send
		GetRemoteData = .ResponseBody
	End With
	Set Retrieval = Nothing
	Set Ads = Server.CreateObject("Adodb.Stream")
	With Ads
		.Type = 1
		.Open
		.Write GetRemoteData
		.SaveToFile server.MapPath(LocalFileName),2
		.Cancel()
		.Close()
	End With
	Set Ads=nothing
end sub

%>
