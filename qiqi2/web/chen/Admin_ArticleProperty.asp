<%@language=vbscript codepage=936 %>
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
<!--#include file="inc/admin_code_article.asp"-->
<%
dim ArticleID,Action,sqlProperty,rsProperty,FoundErr,ErrMsg,PurviewChecked
dim ClassID,tClass,ClassName,RootID,ParentID,Depth,ParentPath,Child,ChildID,tID,tChild,ClassMaster
dim sqlmessage,adname ''退稿插件添加代码
ArticleID=trim(request("ArticleID"))
Action=Trim(Request("Action"))
adname=session("AdminName")
FoundErr=False
if ArticleId="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>请先选定文章！</li>"
end if
if Action="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
end if
if FoundErr=False then
	call SetProperty()
end if
if FoundErr=False then
	'call CloseConn()
        if Action="nopassed" then
	response.Redirect "Admin_ArticleCheck.asp"
	else
	response.Redirect ComeUrl
	end if
else
	'call CloseConn()
	call WriteErrMsg()
end if

sub SetProperty()
	if instr(ArticleID,",")>0 then
		ArticleID=replace(ArticleID," ","")
		sqlProperty="select * from "&jieducm&"_Article where ArticleID in (" & ArticleID & ")"
	else
		ArticleID=Clng(ArticleID)
		sqlProperty="select * from "&jieducm&"_article where ArticleID=" & ArticleID
	end if
	Set rsProperty= Server.CreateObject("ADODB.Recordset")
	rsProperty.open sqlProperty,conn,1,3
	do while not rsProperty.eof
		PurviewChecked=False
		ClassID=rsProperty("ClassID")
		if AdminPurview=1 or AdminPurview_Article<=2 then
			PurviewChecked=True
		else
			if Action="SetPassed" or Action="CancelPassed" then
				set tClass=nt2003.execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,ClassChecker From "&jieducm&"_ArticleClass where ClassID=" & ClassID)
			else
				set tClass=nt2003.execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,ClassMaster From "&jieducm&"_ArticleClass where ClassID=" & ClassID)
			end if
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
				PurviewChecked=CheckClassMaster(tClass(6),AdminName)
				if PurviewChecked=False and ParentID>0 then
					if Action="SetPassed" or Action="CancelPassed" then
						set tClass=nt2003.execute("select ClassChecker from "&jieducm&"_ArticleClass where ClassID in (" & ParentPath & ")")
					else
						set tClass=nt2003.execute("select ClassMaster from "&jieducm&"_ArticleClass where ClassID in (" & ParentPath & ")")
					end if
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
			ErrMsg=ErrMsg & "<br><li>对 " & rsProperty("ArticleID") & " 没有操作权限！</li>"
		else
			select case Action
			case "SetOnTop"
				rsProperty("OnTop") = True
				    '发表加点 （置顶加）  一枝梅添加开始
		          	'if nt2003.site_setting(37)=1 then
			        '   Conn_User.execute("update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "+" & nt2003.site_setting(42) & " where " & db_User_Name & "='" & rsProperty("Editor") & "'")
		           ' end if
					'结束
			case "CancelOnTop"
				rsProperty("OnTop") = False
				    '发表加点 （取消置顶扣）  一枝梅添加开始
		          	'if nt2003.site_setting(37)=1 then
			        '   Conn_User.execute("update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "-" & nt2003.site_setting(42) & " where " & db_User_Name & "='" & rsProperty("Editor") & "'")
		          '  end if
					'结束
			case "SetElite"
				rsProperty("Elite") = True
				    '发表加点 （设推荐加）  一枝梅添加开始
		          '	if nt2003.site_setting(37)=1 then
			       '    Conn_User.execute("update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "+" & nt2003.site_setting(41) & " where " & db_User_Name & "='" & rsProperty("Editor") & "'")
		           ' end if
					'结束
			case "CancelElite"
				rsProperty("Elite") = False
				    '发表加点 （取消推荐扣）  一枝梅添加开始
		          	'if nt2003.site_setting(37)=1 then
			       '    Conn_User.execute("update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "-" & nt2003.site_setting(41) & " where " & db_User_Name & "='" & rsProperty("Editor") & "'")
		           ' end if
					'结束
			case "SetPassed"
				rsProperty("Passed") =True
				'Conn_User.execute("update " & db_User_Table & " set " & db_User_ArticleChecked & "=" & db_User_ArticleChecked & "+1 where " & db_User_Name & "='" & rsProperty("Editor") & "'")
				    '发表加点 （通过审核加）  一枝梅添加开始
		          	'if nt2003.site_setting(37)=1 then
			         '  Conn_User.execute("update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "+" & nt2003.site_setting(38) & " where " & db_User_Name & "='" & rsProperty("Editor") & "'")
		           ' end if
					'结束
				'退稿插件添加代码开始
				'SmsBody="Article_show.asp?ArticleID="&rsProperty("ArticleID")
 				'SmsBody="[url="&SmsBody&"]点击这里查看文章的具体内容[/url]"
				'sqlmessage="insert into "&db_User_Message&" (incept,sender,title,content,sendtime,flag,issend) values ('" & rsProperty("Editor") & "','本站管理员"&adname&"','您发表的文章《"&rsProperty("Title")&"》已经通过审核','"&SmsBody&"',Now(),0,1)"
				'Conn_User.execute(sqlmessage)
				nt2003.delcahe("NewArticle")
			case "nopassed"
				'rsProperty("nopass") =True
				'rsProperty("nopassno") =rsProperty("nopassno")+1
				'rsProperty("nopasstxt") =nopasstxt
				'SmsBody=""&rsProperty("Title")
            	'SmsBody="由于：[color=bb3333]"&nopasstxt&"[/color]，您发表的文章《"&SmsBody&"》未获审核通过，请到后台管理选择被管理员退稿的文章，修改后再重新投稿。"
				'sqlmessage="insert into "&db_User_Message&" (incept,sender,title,content,sendtime,flag,issend) values ('" & rsProperty("Editor") & "','本站管理员"&adname&"','您发表的文章《"&rsProperty("Title")&"》由于下列原因已被管理员第"&rsProperty("nopassno")&"次退回。','"&SmsBody&"',Now(),0,1)"
				'Conn_User.execute(sqlmessage)
					'发表加点 （退稿扣）  一枝梅添加开始
		          	'if nt2003.site_setting(37)=1 then
			        '  Conn_User.execute("update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "-" & nt2003.site_setting(39) & " where " & db_User_Name & "='" & rsProperty("Editor") & "'")
		            'end if
					'结束
			case "deleted"
					'发表加点 （删除扣）  一枝梅添加开始
		          '	if nt2003.site_setting(37)=1 then
			         '  Conn_User.execute("update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "-" & nt2003.site_setting(40) & " where " & db_User_Name & "='" & rsProperty("Editor") & "'")
		          '  end if
					'结束
				'SmsBody=""&rsProperty("Title") 	
				'SmsBody="您发表的文章《"&SmsBody&"》未获审核通过，已经删除，仅代表网站全体管理人员向您表示十分遗憾。"
				'sqlmessage="insert into "&db_User_Message&" (incept,sender,title,content,sendtime,flag,issend) values ('" & rsProperty("Editor") & "','本站管理员"&adname&"','您发表的文章《"&rsProperty("Title")&"》未通过审核，已经删除','"&SmsBody&"',Now(),0,1)"
				'Conn_User.execute(sqlmessage)
				conn.execute("delete from "&jieducm&"_Article where ArticleID in (" & ArticleID & ")")
				conn.execute("delete from "&jieducm&"_ArticleComment where ArticleID in (" & ArticleID & ")")
				'退稿插件添加代码结束
			case "CancelPassed"
				rsProPerty("Passed") =False
				'Conn_User.execute("update " & db_User_Table & " set " & db_User_ArticleChecked & "=" & db_User_ArticleChecked & "-1 where " & db_User_Name & "='" & rsProperty("Editor") & "'")
					'发贴加点 （取消审核扣）  一枝梅添加开始
		          	'if nt2003.site_setting(37)=1 then
			          ' Conn_User.execute("update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "-" & nt2003.site_setting(38) & " where " & db_User_Name & "='" & rsProperty("Editor") & "'")
		           ' end if
					'结束
			end select
			rsProperty.update
		end if
		rsProperty.movenext
	loop
	rsProperty.close
	set rsProperty=nothing
end sub
%>