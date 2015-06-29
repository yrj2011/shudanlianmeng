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
Const PurviewLevel=2
Const CheckChannelID=2
Const PurviewLevel_Article=3
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/admin_code_Article.asp"-->
<%
dim ArticleID,Action,sqlDel,rsDel,FoundErr,ErrMsg,PurviewChecked,ObjInstalled
dim ClassID,tClass,ClassName,RootID,ParentID,Depth,ParentPath,Child,ChildID,tID,tChild,ClassMaster
ArticleID=trim(request("ArticleID"))
Action=Trim(Request("Action"))
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
FoundErr=False
if Action="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
end if
if FoundErr=False then
	if Action="Del" then
		call DelArticle()
	elseif Action="ConfirmDel" then
		call ConfirmDel()
	elseif Action="ClearRecyclebin" then
		call ClearRecyclebin()
	elseif Action="Restore" then
		call Restore()
	elseif Action="RestoreAll" then
		call RestoreAll()
	elseif Action="DelFromSpecial" then
		call DelFromSpecial()
	end if
end if
if FoundErr=False then
	'call CloseConn()
	response.Redirect ComeUrl
else
	call WriteErrMsg()
	'call CloseConn()
end if

sub DelArticle()
	if ArticleID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请先选定文章！</li>"
		exit sub
	end if

	if instr(ArticleID,",")>0 then
		ArticleID=replace(ArticleID," ","")
		sqlDel="select * from "&jieducm&"_Article where ArticleID in (" & ArticleID & ")"
	else
		ArticleID=Clng(ArticleID)
		sqlDel="select * from "&jieducm&"_article where ArticleID=" & ArticleID
	end if
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	do while not rsDel.eof
		
		PurviewChecked=False
		ClassID=rsDel("ClassID")
		if AdminPurview=1 or AdminPurview_Article<=2 or (rsDel("Editor")=AdminName and rsDel("Passed")=False) then
			PurviewChecked=True
		else
				set tClass=nt2003.execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,ClassMaster From "&jieducm&"_ArticleClass where ClassID=" & ClassID)
				if tClass.bof and tClass.eof then
					founderr=True
					ErrMsg=ErrMsg & "<br><li>找不到指定的栏目</li>"
				else
					ClassName=tClass(0)
					RootID=tClass(1)
					ParentID=tClass(2)
					Depth=tClass(3)
					ParentPath=tClass(4)
					Child=tClass(5)
					ClassMaster=tClass(6)
					PurviewChecked=CheckClassMaster(ClassMaster,AdminName)
					if PurviewChecked=False and ParentID>0 then
						set tClass=nt2003.execute("select ClassMaster from "&jieducm&"_ArticleClass where ClassID in (" & ParentPath & ")")
						do while not tClass.eof
							PurviewChecked=CheckClassMaster(tClass(0),AdminName)
							if PurviewChecked=True then exit do
							tClass.movenext
						loop
					end if
				end if
		end if
		if PurviewChecked=False then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>删除" & rsDel("ArticleID") & "失败！原因：没有操作权限！</li>"
		else
			rsDel("Deleted")=True
			rsDel.update
		end if
		rsDel.movenext
	loop
	rsDel.close
	set rsDel=nothing
end sub

sub ConfirmDel()
	if AdminPurview=2 and AdminPurview_Article>1 then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>对不起，你的权限不够！</li>"
		exit sub
	end if
	if ArticleID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请先选定文章！</li>"
		exit sub
	end if
	ArticleID=replace(ArticleID," ","")
	sqlDel="select UploadFiles from "&jieducm&"_article where ArticleID in (" & ArticleID & ")"
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	do while not rsDel.eof
		call DelFiles(rsDel(0) & "")
		rsDel.movenext
	loop
	rsDel.close
	set rsDel=nothing
	'===================== 山风签收插件修改01开始 =====================
	call delReceive()
	'--------------------- 山风签收插件修改01结束 ---------------------
	dim dsql,drs,fso,filepath
	IF instr(ArticleID,",") >0 Then		
		dsql = "select newspath from "&jieducm&"_Article where ArticleID in ("&ArticleID&")"
		Set drs = Server.CreateObject("Adodb.RecordSet")
		drs.Open dsql,conn,1,3
		Do While Not drs.Eof
			Set fso = Server.CreateObject("Scripting.FileSystemObject")
			filepath = Server.MapPath("../news/"&drs("newspath"))
			IF fso.fileexists(filepath) then
			
			End IF
			drs.delete
		drs.MoveNext
		loop
	Else
		dsql = "select newspath from "&jieducm&"_Article where ArticleID="&ArticleID
		Set drs = Server.CreateObject("Adodb.RecordSet")
		drs.Open dsql,conn,1,3
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		filepath = Server.MapPath("../news/"&drs("newspath"))
		IF fso.fileexists(filepath) then
			fso.Deletefile (filepath)
		End IF
		drs.delete
	End IF
	'conn.execute("delete from Article where ArticleID in (" & ArticleID & ")")
	'conn.execute("delete from ArticleComment where ArticleID in (" & ArticleID & ")")
end sub

sub ClearRecyclebin()
	if AdminPurview=2 and AdminPurview_Article>1 then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>对不起，你的权限不够！</li>"
		exit sub
	end if
	ArticleID=""
	sqlDel="select ArticleID,UploadFiles from "&jieducm&"_article where Deleted=1"
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	do while not rsDel.eof
		if ArticleID="" then
			ArticleID=rsDel(0)
		else
			ArticleID=ArticleID & "," & rsDel(0)
		end if
		call DelFiles(rsDel(1) & "")
		rsDel.movenext
	loop
	rsDel.close
	set rsDel=nothing
	if ArticleID<>"" then
		conn.execute("delete from "&jieducm&"_Article where Deleted=1")
		conn.execute("delete from "&jieducm&"_ArticleComment where ArticleID in (" & ArticleID & ")")
	end if
end sub

sub Restore()
	if AdminPurview=2 and AdminPurview_Article>1 then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>对不起，你的权限不够！</li>"
		exit sub
	end if
	if ArticleID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请先选定文章！</li>"
		exit sub
	end if
	ArticleID=replace(ArticleID," ","")
	conn.execute("update "&jieducm&"_Article set Deleted=0 where ArticleID in (" & ArticleID & ")")
	if instr(ArticleID,",")>0 then
		ArticleID=replace(ArticleID," ","")
		sqlDel="select * from "&jieducm&"_Article where ArticleID in (" & ArticleID & ")"
	else
		ArticleID=Clng(ArticleID)
		sqlDel="select * from "&jieducm&"_article where ArticleID=" & ArticleID
	end if
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	rsDel.close
	set rsDel=nothing

end sub

sub RestoreAll()
	if AdminPurview=2 and AdminPurview_Article>1 then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>对不起，你的权限不够！</li>"
		exit sub
	end if
	sqlDel="select * from "&jieducm&"_Article where Deleted=1"
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	rsDel.close
	set rsDel=nothing
end sub

sub DelFromSpecial()
	if AdminPurview=2 and AdminPurview_Article>1 then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>对不起，你的权限不够！</li>"
		exit sub
	end if
	ArticleID=replace(ArticleID," ","")
	conn.execute("update "&jieducm&"_Article set SpecialID=0 where ArticleID in (" & ArticleID & ")")
end sub


'===================== 山风签收插件修改02开始 =====================
sub delReceive()

	conn.execute("delete from "&jieducm&"_ArticleReceive where ArticleID in (" & ArticleID & ")")
end sub
'--------------------- 山风签收插件修改02结束 ---------------------
%>
