<%@language=vbscript codepage=936 %>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/md5.asp"-->

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
Class System_Cls
	Private LocalCacheName,Cache_Data
	Public Reloadtime,CacheName,CacheData,savelog,SqlQueryNum '新增变量
	Public pNum,pNum2

	'声明System_Cls类预处理内容
	Private Sub Class_Initialize()
		Dim UserAgent,web_CacheName
		web_CacheName = "Jxqgw"   '缓存名称，如果一个站点有多个站请更改成不同名称
		UserAgent = Trim(Lcase(Request.Servervariables("HTTP_USER_AGENT")))
		If InStr(UserAgent,"teleport") > 0 or InStr(UserAgent,"webzip") > 0 or InStr(UserAgent,"flashget")>0 or InStr(UserAgent,"offline")>0 Then
			Response.Write "请不要采用teleport/Webzip/Flashget/Offline等工具来浏览网站！"
			Response.End
		End If
		CacheName=Replace(Server.MapPath("\index.asp"),"index.asp","")
		'Response.Write(CacheName)
		'ssssResponse.End()
		if right(CacheName,3)="inc" then
			CacheName=Replace(CacheName,"inc","")
		end if
		CacheName=Replace(CacheName,":","")
		CacheName=Replace(CacheName,"\","")	'重大错误，发现修正	
		CacheName=CacheName & web_CacheName  '添加
		Reloadtime=14400
		savelog=0
		SqlQueryNum=0
		pNum=1:pNum2=0
	End Sub

	'声明System_Cls类终止处理内容
	Private Sub class_terminate()
		If IsObject(Conn) Then 
			Conn.Close
			Set Conn = Nothing
		End If 
	End Sub

	'Cache处理过程
	Public Property Let Name(ByVal vNewValue)
		LocalCacheName=LCase(vNewValue)
	End Property

	Public Property Let Value(ByVal vNewValue)
		If LocalCacheName<>"" Then 
			ReDim Cache_Data(2)
			Cache_Data(0)=vNewValue
			Cache_Data(1)=Now()
			Application.Lock
			Application(CacheName & "_" & LocalCacheName) = Cache_Data
			Application.unLock
		Else
			Err.Raise vbObjectError + 1, "CacheServer", " please change the CacheName."
		End If
	End Property
	Public Property Get Value()
		If LocalCacheName<>"" Then 
			Cache_Data=Application(CacheName & "_" & LocalCacheName)	
			If IsArray(Cache_Data) Then
				Value=Cache_Data(0)
			Else
				Err.Raise vbObjectError + 1, "CacheServer", " The Cache_Data("&LocalCacheName&") Is Empty."
			End If
		Else
			Err.Raise vbObjectError + 1, "CacheServer", " please change the CacheName."
		End If
	End Property
	Public Function ObjIsEmpty()
		ObjIsEmpty=True	
		Cache_Data=Application(CacheName & "_" & LocalCacheName)
		If Not IsArray(Cache_Data) Then Exit Function
		If Not IsDate(Cache_Data(1)) Then Exit Function
		If DateDiff("s",CDate(Cache_Data(1)),Now()) < (60*Reloadtime) Then ObjIsEmpty=False		
	End Function
	Public Sub DelCahe(MyCaheName)
		Application.Lock
		Application.Contents.Remove(CacheName&"_"&MyCaheName)
		Application.unLock
	End Sub



	'定义系统资源变量
	Public Site_Info,Site_Setting,Site_Version,Site_Copyright,BadWords,rBadWord
	'取得系统定义资源
	Public Sub GetSite_Setting()
		Name="setup"
		If ObjIsEmpty() Then ReloadSetup()
		CacheData=value

		'每日更新数据
		Name="Date"
		'第一次起用网站或者重启IIS的时候加载缓存
		If ObjIsEmpty() Then
			value=Date()
		End If
		Name="Date"
		If Cstr(value) <> Cstr(Date()) Then
			Name="setup"
		        value=Date()
			ReloadSetup()
			CacheData=value
			DelCahe("SiteCount")
		End If
		Dim Setting
		Setting = CacheData(1,0)
				
		Dim SQL,Rs,i
		SQL = "Select top 1 * from [AC_setup]"
		Set Rs = Execute(SQL)
		Setting = Rs("Site_Setting")
		Set Rs = Nothing
		
		Setting = Split(Setting,"|||")
		Site_Info = Setting(0)
		Site_Info = Split(Site_Info,",")
		Site_Setting = Setting(1)
		Site_Setting = Split (Site_Setting,",")
		Site_Version = CacheData(2,0)
		Site_Copyright = CacheData(3,0)
		BadWords = Split(CacheData(5,0),"|")
		rBadWord = Split(CacheData(6,0),"|")
	End Sub

	Public Sub ReloadSetup()
		Dim SQL,Rs,i
		SQL = "Select * from [AC_setup]"
		Set Rs = Execute(SQL)
		value = Rs.GetRows(1)
		Set Rs = Nothing
	End Sub 


  '定义风格相关变量
	  
		 

	 
  '已经删除

	'页面显示类函数
	  '已经删除


'页面显示内容函数
	  '已经删除

	'读取网站访问记数
    'DELETE

	Public sub loadVote()
		dim Rs,SQL,tmpdata,i
		Name="vote"&ChannelID
		SQL="select top 1 * from Vote where IsSelected=True and (ChannelID=0 or ChannelID=" & ChannelID & ") order by ID Desc"
		Set Rs= execute(SQL)
		if Rs.bof and Rs.eof then 
			tmpdata = "<font color='#ff9900'>·&nbsp;</font>没有任何调查" 
		else
			tmpdata = "<form name='VoteForm' method='post' action='vote.asp' target='_blank'>"
			tmpdata = tmpdata & "&nbsp;&nbsp;&nbsp;&nbsp;" & Rs("Title") & "<br>"
			if Rs("VoteType")="Single" then
				for i=1 to 8
					if trim(Rs("Select" & i) & "")="" then exit for
					tmpdata = tmpdata & "<input type='radio' name='VoteOption' value='" & i & "' style='border:0'>" & Rs("Select" & i) & "<br>"
				next
			else
				for i=1 to 8
					if trim(Rs("Select" & i) & "")="" then exit for
					tmpdata = tmpdata & "<input type='checkbox' name='VoteOption' value='" & i & "' style='border:0'>" & Rs("Select" & i) & "<br>"
				next
			end if
			tmpdata = tmpdata & "<br><input name='VoteType' type='hidden'value='" & Rs("VoteType") & "'>"
			tmpdata = tmpdata & "<input name='Action' type='hidden' value='Vote'>"
			tmpdata = tmpdata & "<input name='ID' type='hidden' value='" & Rs("ID") & "'></form>" 
			tmpdata = tmpdata & "<div align='center'>"
			tmpdata = tmpdata & "<a href='javascript:VoteForm.submit();'><img src='Skin/51dsn03/bt_voteSubmit.gif' border='0'></a>" 
	        	tmpdata = tmpdata & "<a href='Vote.asp?ID=" & Rs("ID") & "&Action=Show' target='_blank'><img src='Skin/51dsn03/bt_voteView.gif' border='0'></a>" 
			tmpdata = tmpdata & "</div>" 
		end if
		Rs.close
		set Rs=nothing
		value=tmpdata
	end sub

'*************************************************
'缓存文章栏目 '不知道是前台用的还是
'=================================================
'过程名：ArticleContentshiyu
'作  用：显示文章属性、标题、作者、更新日期、点击数等信息
'参  数：intTitleLen  ----标题最多字符数，一个汉字=两个英文字符
'        ShowProperty ----是否显示文章属性（固顶/推荐/普通），True为显示，False为不显示
'        ShowIncludePic ---是否显示“[图文]”字样，True为显示，False为不显示
'        ShowAuthor -------是否显示文章作者，True为显示，False为不显示
'        ShowDateType -----显示更新日期的样式，0为不显示，1为显示年月日，2为只显示月日。
'        ShowHits ---------是否显示文章点击数，True为显示，False为不显示
'        ShowHot ----------是否显示热门文章标志，True为显示，False为不显示
'=================================================
function ArticleContentshiyu(intTitleLen,ShowProperty,ShowIncludePic,ShowAuthor,ShowDateType,ShowHits,ShowHot)
   	dim i,strTemp,TitleStr,Author,AuthorName,AuthorEmail
    	i=0
	strTemp="<table width=100%  border=0 cellpadding=0 cellspacing=0>"
	do while not rsArticle.eof
		strTemp=strTemp & "<tr><td>"
		if ShowProperty=True then
			if rsArticle("OnTop")=true then
				strTemp = strTemp & "<img src='images/article_ontop.gif' alt='固顶文章'>&nbsp;"
			elseif rsArticle("Elite")=true then
				strTemp = strTemp & "<img src='images/article_elite.gif' alt='推荐文章'>&nbsp;"
			else
				strTemp = strTemp & "<img src='images/article_common.gif' alt='普通文章'>&nbsp;"
			end if
		end if
		if ShowIncludePic=True and rsArticle("IncludePic")=true then
			strTemp = strTemp & "<font color=blue>[图文]</font>"
		end if
		Author=rsArticle("Author")
		if instr(Author,"|")>0 then
			AuthorName=left(Author,instr(Author,"|")-1)
			AuthorEmail=right(Author,len(Author)-instr(Author,"|")-1)
		else
			AuthorName=Author
			AuthorEmail=""
		end if
		strTemp = strTemp & "<a href='" & rsArticle("LayoutFileName") & "?ArticleID=" & rsArticle("articleid") & "' title='文章标题：" & rsArticle("Title") & vbcrlf & "文章作者：" & AuthorName & vbcrlf & "更新时间：" & rsArticle("UpdateTime") & vbcrlf & "点击次数：" & rsArticle("Hits") & "' target='_blank'>"
		TitleStr=gotTopic(rsArticle("title"),intTitleLen)
		if rsArticle("TitleFontType")=1 then
			TitleStr="<b>" & TitleStr & "</b>"
		elseif rsArticle("TitleFontType")=2 then
			TitleStr="<em>" & TitleStr & "</em>"
		elseif rsArticle("TitleFontType")=3 then
			TitleStr="<b><em>" & TitleStr & "</em></b>"
		end if
		if rsArticle("TitleFontColor")<>"" then
			TitleStr="<font color='" & rsArticle("TitleFontColor") & "'>" & TitleStr & "</font>"
		end if
		strTemp=strTemp & TitleStr & "</a></td><td align=right>"
		if ShowAuthor=True or ShowDateType>0 or ShowHits=True then
			strTemp = strTemp & "["
			if ShowAuthor=True then
				if AuthorEmail="" then
					strTemp=strTemp & "<font color=#999999>" & AuthorName & "</font>"
				else
					strTemp=strTemp & "<font color=#999999><a href='mailto:" & AuthorEmail & "'>" & AuthorName & "</a></font>"
				end if
			end if
			if ShowDateType>0 then
				if ShowAuthor=True then
					strTemp=strTemp & "|"
				end if
				if CDate(FormatDateTime(rsArticle("UpdateTime"),2))=date() then
					strTemp = strTemp & "<font color=red>"
				else
					strTemp= strTemp & "<font color=#999999>"
				end if
				if ShowDateType=1 then
				  if month(rsArticle("UpdateTime"))<10 then
				    strTemp= strTemp & "0"
				  end if 
					strTemp= strTemp & month(rsArticle("UpdateTime"))& "-"
				  if day(rsArticle("UpdateTime"))<10 then
				    strTemp= strTemp & "0"
				  end if
				    strTemp= strTemp & day(rsArticle("UpdateTime")) & "</font>"
				else
					strTemp=strTemp & FormatDateTime(rsArticle("UpdateTime"),1) & "</font>"
				end if
			end if
			if ShowHits=True then
				if ShowAuthor=True or ShowDateType>0 then
					strTemp=strTemp & "|"
				end if
				strTemp=strTemp & "<font color=#999999>" & rsArticle("Hits") & "</font>"
			end if
			strTemp=strTemp  & "]"
		end if
		if ShowHot=True and rsArticle("Hits")>=nt2003.site_setting(14) then
			strTemp= strTemp & "<img src='images/hot.gif' alt='热点文章'>"
		end if
		strTemp= strTemp & "</td></tr>"
		rsArticle.movenext
		i=i+1
		if i>=MaxPerPage then exit do
	loop
	strTemp=strTemp & "</table>"
	ArticleContentshiyu=strTemp
end function

public tmpdata
public function showarticleshiyu()
	dim sqlRoot,rsRoot
	name="classroot"
	If ObjIsEmpty() Then
		sqlRoot="select C.ClassID,C.ClassName,C.RootID,L.LayoutFileName,L.LayoutID,C.Child From "&jieducm&"_ArticleClass C inner join Layout L on C.LayoutID=L.LayoutID where C.ParentID=0 and IsElite=True and LinkUrl='' order by C.RootID"
		Set rsRoot = Execute(sqlRoot)
		if rsRoot.bof and rsRoot.eof then
			value=""
		else
			value = rsRoot.GetString(,,"|||","@@@","")
		end if
		rsRoot.close
		set rsRoot=nothing
	end if
	tmpdata=value

	Name="showarticleshiyu"
	If ObjIsEmpty() Then loadShowarticleshiyu()
	showarticleshiyu=value
end function

public sub loadShowarticleshiyu()
	dim strtemp
	strtemp = ""
	dim trs,arrClassID,TitleStr,iClassID,strrow,strcol,i

	dim sqlClassAD,rsClassAD,ClassAD
	sqlClassAD="select * from Advertisement where IsSelected=True and (ChannelID=0 or ChannelID=" & ChannelID & ") and ADType=2 order by ID Desc"
	Set rsClassAD=Execute(sqlClassAD)
	if tmpdata="" then 
		strtemp = strtemp & "<tr><td height='60' class='tdbg_mainall_lanmu'><font color='#ff9900'>·&nbsp;</font>"
		strtemp = strtemp & "还没有任何栏目，请首先添加栏目。"
		strtemp = strtemp & "</td></tr>"
	else
		strtemp = strtemp & "<tr><td class='tdbg_mainall_lanmu'><table width='100%'  border='0' cellspacing='0' cellpadding='0'></tr>"
		strrow=Split(tmpdata,"@@@")
		iClassID=0
		for i = 0 to UBound(strrow)-1
			strcol=Split(strrow(i),"|||")
			strtemp = strtemp & "<td valign='top' width='49%'><table width='100%'  border='0' cellspacing='0' cellpadding='0'>"
			strtemp = strtemp & "<tr><td class='title_right'><table width='100%'  border='0' cellspacing='0' cellpadding='0'>"
			strtemp = strtemp & "<tr><td width='14'><img src='{$PicUrl}/h_cl1.gif' width='14' height='23'></td>"
			strtemp = strtemp & "<td>"
			arrClassID=strcol(0)
			strtemp = strtemp & "<a href='" & strcol(3) & "?ClassID=" & strcol(0) & "'><strong>" & strcol(1) & "</strong></a>"
			if strcol(5)>0 then
			set trs=execute("select ClassID from "&jieducm&"_ArticleClass where RootID=" & strcol(2) & " and Child=0 and LinkUrl=''")
				do while not trs.eof
					arrClassID=arrClassID & "," & trs(0)
					trs.movenext
				loop
			end if
			strtemp = strtemp & "</td><td width='60' align='right'>"
			strtemp = strtemp & "<a href='" & strcol(3) & "?ClassID=" & strcol(0) & "'>more...</a>&nbsp;"
			strtemp = strtemp & "</td>"
			strtemp = strtemp & "</tr>"
			strtemp = strtemp & "</table></td>"
			strtemp = strtemp & "</tr>"
			strtemp = strtemp & "<tr><td height='127' valign='top' class='tdbg_right'>"

			sql="select top 6 A.ArticleID,A.ClassID,L.LayoutID,L.LayoutFileName,A.Title,A.Key,A.Author,A.CopyFrom,A.UpdateTime,A.Editor,A.TitleFontColor,A.TitleFontType,"
			sql=sql & "A.Hits,A.OnTop,A.Hot,A.Elite,A.Passed,A.IncludePic,A.Stars,A.PaginationType,A.ReadLevel,A.ReadPoint,A.DefaultPicUrl from Article A"
			sql=sql & " inner join Layout L on A.LayoutID=L.LayoutID where A.Deleted=False and A.Passed=True and A.ClassID in (" & arrClassID & ")  order by A.OnTop,A.ArticleID desc"
			set rsArticle=execute(sql)
			if rsArticle.bof and  rsArticle.eof then
				strtemp = strtemp & "<font color='#ff9900'>·&nbsp;</font>此栏目没有任何文章"
			else
				strtemp = strtemp & ArticleContentshiyu(22,True,True,False,1,False,False)
			end if
			rsArticle.close

			strtemp = strtemp & "</td>"
			strtemp = strtemp & "</tr>"
			strtemp = strtemp & "</table></td>"
			iClassID=iClassID+1
			if iClassID mod 2=0 then
				strtemp = strtemp & "</tr>"
				if not rsClassAD.bof and not rsClassAD.eof then
					if rsClassAD("isflash")=true then
						ClassAD= "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0'"
						if rsClassAD("ImgWidth")>0 then ClassAD = ClassAD & " width='" & rsClassAD("ImgWidth") & "'"
						if rsClassAD("ImgHeight")>0 then ClassAD = ClassAD & " height='" & rsClassAD("ImgHeight") & "'"
						ClassAD = ClassAD & "><param name='movie' value='" & rsClassAD("ImgUrl") & "'><param name='quality' value='high'><embed src='" & rsClassAD("ImgUrl") & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'"
						if rsClassAD("ImgWidth")>0 then ClassAD = ClassAD & " width='" & rsClassAD("ImgWidth") & "'"
						if rsClassAD("ImgHeight")>0 then ClassAD = ClassAD & " height='" & rsClassAD("ImgHeight") & "'"
						ClassAD = ClassAD & "></embed></object>"
					else
						ClassAD ="<a href='" & rsClassAD("SiteUrl") & "' target='_blank' title='" & rsClassAD("SiteName") & "：" & rsClassAD("SiteUrl") & vbcrlf & rsClassAD("SiteIntro") & "'><img src='" & rsClassAD("ImgUrl") & "'"
						if rsClassAD("ImgWidth")>0 then ClassAD = ClassAD & " width='" & rsClassAD("ImgWidth") & "'"
						if rsClassAD("ImgHeight")>0 then ClassAD = ClassAD & " height='" & rsClassAD("ImgHeight") & "'"
						ClassAD = ClassAD & " border='0'></a>"
					end if
					strtemp = strtemp & "<tr><td align='center' bgcolor='#E4EEFD' class='tdbg_mainall' colSpan='3'>"
					strtemp = strtemp & ClassAD
					strtemp = strtemp & "</td></tr>"
					rsClassAD.movenext
				end if
				strtemp = strtemp & "</tr><tr><td height='6'></td></tr>"
			else
				strtemp = strtemp & "<td width='1%'></td>"
			end if
		next
	end if
	strtemp = strtemp & "</table></td></tr>"
	value = strtemp
end sub
'以上该程序可以删除


'其他处理函数(安全，字符过滤等)
	Public Function Execute(Command)
		If Not IsObject(Conn) Then ConnectionDatabase
		If IsDeBug = 0 Then 
			On Error Resume Next
			Set Execute = Conn.Execute(Command)
			If Err Then
				err.Clear
				Set Conn = Nothing
				Response.Write "查询数据的时候发现错误，请检查您的查询代码是否正确。"
				Response.End
			End If
		Else
			'Response.Write(Request.ServerVariables("URL"))
			Set Execute = Conn.Execute(Command)
		End If	
		SqlQueryNum = SqlQueryNum+1
	End Function

	Public Function strLength(str)
		If isNull(str) Or Str = "" Then
			StrLength = 0
			Exit Function
		End If
		Dim WINNT_CHINESE
		WINNT_CHINESE=(len("例子")=2)
		If WINNT_CHINESE Then
			Dim l,t,c
			Dim i
			l=len(str)
			t=l
			For i=1 To l
				c=asc(mid(str,i,1))
				If c<0 Then c=c+65536
				If c>255 Then t=t+1
			Next
			strLength=t
		Else 
			strLength=len(str)
		End If
	End Function
	Public Function ChkBadWords(Str)
		If IsNull(Str) Then Exit Function
		Dim i
		For i = 0 To Ubound(BadWords)
			If i > UBound(rBadWord) Then
				Str = Replace(Str,BadWords(i),"*")
			Else
				Str = Replace(Str,BadWords(i),rBadWord(i))
			End If
		Next
		ChkBadWords = Str
	End Function
	Public Function Checkstr(Str)
		dim tempcheckstr
		tempcheckstr=str
		If Isnull(Str) Then
			CheckStr = ""
			Exit Function 
		End If
		CheckStr = Replace(tempcheckstr,"'","''")
	End Function
End Class

nt2003.GetSite_Setting() 
dim sql,rs,rs_jieducm
'dim username,password,CheckCode
username=request.form("username")
password=request.form("password")
username=replace(username,"'","")
username=replace(username,"""","")
username=replace(username,"=","")
password=replace(password,"'","")
password=replace(password,"""","")
password=replace(password,"=","")
CheckCode=replace(trim(Request("CheckCode")),"'","")
if UserName="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>用户名不能为空！</li>"
end if
if Password="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>密码不能为空！</li>"
end if
if CheckCode="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>验证码不能为空！</li>"
end if
if session("CheckCode")="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>你登录时间过长，请重新返回登录页面进行登录。</li>"
end if
if CheckCode<>CStr(session("CheckCode")) then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>您输入的确认码和系统产生的不一致，请重新输入。</li>"
end if
if FoundErr<>True then
	password=md5(password)
	sql="select * from "&jieducm&"_admin where password='"&password&"' and username='"&username&"'"
	set rs=nt2003.execute(sql) 
	if rs.bof and rs.eof then
		Set rs_jieducm=server.createobject("ADODB.RECORDSET")
	    rs_jieducm.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rs_jieducm.addnew
		rs_jieducm("username")=username
		rs_jieducm("class")="登录后台"
		rs_jieducm("content")=username&"登录时间:"&now&"登录IP:"&Request.ServerVariables("REMOTE_ADDR")&""
		rs_jieducm("jiegou")="失败"
    	rs_jieducm.update
    	rs_jieducm.close
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>用户名或密码错误！！！你本次登陆的ip:"&Request.ServerVariables("REMOTE_ADDR")&"已记录在案</li>"
		
	else
		if password<>rs("password") then
		Set rs_jieducm=server.createobject("ADODB.RECORDSET")
	    rs_jieducm.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rs_jieducm.addnew
		rs_jieducm("username")=username
		rs_jieducm("class")="登录后台"
		rs_jieducm("content")=username&"登录时间:"&now&"登录IP:"&Request.ServerVariables("REMOTE_ADDR")&""
		rs_jieducm("jiegou")="失败"
    	rs_jieducm.update
    	rs_jieducm.close
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>用户名或密码错误！！你本次登陆的ip:"&Request.ServerVariables("REMOTE_ADDR")&"已记录在案！</li>"
		else
		 
			nt2003.execute("update "&jieducm&"_Admin set LastLoginIP='"&Request.ServerVariables("REMOTE_ADDR")&"',LastLoginTime='"&now&"',LoginTimes=LoginTimes+1 Where username='"&username&"'") 
			session.Timeout=nt2003.site_setting(15)
			session("AdminName")=username
		
		Set rs_jieducm=server.createobject("ADODB.RECORDSET")
	    rs_jieducm.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rs_jieducm.addnew
		rs_jieducm("username")=username
		rs_jieducm("class")="登录后台"
		rs_jieducm("content")=username&"登录时间:"&now&"登录IP:"&Request.ServerVariables("REMOTE_ADDR")&""
		rs_jieducm("jiegou")="成功"
    	rs_jieducm.update
    	rs_jieducm.close
		
			rs.close
			set rs=nothing
			Response.Redirect "Admin_Index.asp"
		end if
	end if
	rs.close
	set rs=nothing
end if
if FoundErr=True then
	call WriteErrMsg()
end if
'call CloseConn()

'****************************************************
'过程名：WriteErrMsg
'作  用：显示错误提示信息
'参  数：无
'****************************************************
sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>错误信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td height='22' class='title'><strong>错误信息</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr><td height='100' class='tdbg' valign='top'><b>产生错误的可能原因：</b><br>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td class='tdbg'><a href='Admin_Login.asp'>&lt;&lt; 返回登录页面</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub
%>