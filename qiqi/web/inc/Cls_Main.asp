<%

Class Cls_YSvoidMain
  public web_vt,web_name,web_url,Web_RealURL,web_email,web_now_dir,web_urls,web_dir,web_cookies,web_skin,web_dir_skin,web_unit,now_time,today_time,timer_start,pro_edition,num_rs

	Public YSvoid_Info			'系统一些统计信息
	Public Web_Common
	Public IsHealth
	Public SuperAdmins, AdminFolder, Web_Upload, Web_UploadFilters, Web_BadWords, Web_LockIP
	Public Web_LockTime, Web_Boy, Web_Girl, Web_Index
	Public Web_AdminCodes ,Web_Number ,Web_UploadType ,Web_MailServer ,Web_RegBadWords, Web_RegUserName
	Public Web_CopyRight, Web_Closer, Web_Mode
	Public Web_UserLevel, Web_PerPay, Web_Medal, Web_Gbook, Web_ForumData, Web_Payment
	Public IsDeBug
	Public Reloadtime,CacheName
	Public Fso_Sys_Var										'FSO 对象名
	Private LocalCacheName,Cache_Data

	Public IsConn											'Conn 是否开启

	Private WebConfig										'网站常规设置，数组
	'模板私有变量
	Private iMainSkinID
	Private iMainName							'当前模板名称
	Private iMainIndex							'首页模板
	Private iMainWhole							'模板框架
	Private iMainOther, iMainOtherName			'其他模板、标签名称
	Private iMainStrer, iMainStrerName			'全局字符串、标签名称
	Private iPageStrer, iPageStrerName			'频道字符串、标签名称
	Private iPageOther, iPageOtherName			'频道其他模板、标签名称
	Private iMainWidth							'宽度值
	Private iMainPicPath						'图片路径
	Private iMainCssPath						'Css路径
	Private iMainConfig							'框架判断
	Private PutHtml

	Public total_counter									'统计流量
	Public IsView'记得要删除的
	Public PageModTrue										'判断栏目是否设置模板

	Public ChannelTrue
	Public ChannelName,ChannelTit,ChannelDir,ChannelTable,ChannelUnit,ChannelListNum,ChannelPosition
	Public ChannelSetup,ClassDepth,ChannelPutType,ChannelCuteNum,ChannelLeftCuteNum,ChannelCount,ChannelHits,ChannelOption,ChannelReward,ChannelDates,ChannelRss,ChannelPower,ChannelPHidden
	Public ChannelUpType,ChannelUpMax,ChannelUpDayNum,ChannelUpFileType,ChannelUploadDir
	Public ChannelPicWidth, ChannelPicHeight
	Public ReviewTrue, SpecialTrue, KeyWordTrue, ErrTrue, CastTrue
	Public ChannelSpecial, ChannelKeyWord, ChannelCast

	'//* 类初始化 *//
	Private Sub Class_Initialize()
		If Not Response.IsClientConnected Then Response.End
		IsDeBug = 1
		IsView = 0
		Reloadtime = 14400
		Web_cookies = InstallCookies
		total_counter=0
		YSvoid_Online=0
		YSvoid_UserOnline=0
		YSvoid_GuestOnline=0
		IsConn = False
		ChannelTrue = False
		ReviewTrue = False
		SpecialTrue = False
		CastTrue = False
		KeyWordTrue = False
		ErrTrue = False
		IsHealth=False
		now_time=time_type(now(),1)
		today_time=time_type(now(),4)
		timer_start=timer()
		pro_edition="Y"&"S"&"v"&"o"&"i"&"d"&" C"&"M"&"S"&" V"&"5"&"."&"0"&" S"&"P"&"1"
		num_rs=0
		call own_dir()
	End Sub

    '//* 类结束 *//
    Private Sub Class_Terminate()

    End Sub

	Public Function Format_Mid_Num(ByVal dNum)
	Format_Mid_Num=YSvoid.Web_Number(dNum)
    End Function

    Public Function Web_Dim(ByVal dNum)
	Web_Dim=YSvoid.Web_Common(dNum)
    End Function

	Public Sub Cms_InitLoad()	
		Cms_InfoInit()
		Web_Name = Web_Common(1)
		Web_Url = Web_Common(2)
		Web_Unit = Web_Common(8)
		Web_Boy = Web_Common(9)
		Web_Girl = Web_Common(10)
		Web_Index = Web_Common(11)
		Web_Mode = Int(Web_Common(12))
		MainMod_Load()
		Online_Init()
	End Sub


	'//* 载入系统常规配置 *//
	Private Sub Cms_InfoInit()
		Dim Tmpstr
		'0-注册人数、1-最后注册用户、2-访问人数、3-最高在线人数，4-最高访问时间、5-网站开始时间
		'6-今日时间、7-今日新贴、8-配置参数、9-超级用户、10-网站Fso、11-后台目录、12-上传目录
		'13-不健康字符、14-锁定IP、15-用户等级、16-预付费、17-留言参数、18-当前论坛数据表信息
		'19-支付参数、20-上传过滤
		Name = "YSvoid_info"
		If Cache_Chk() Then
			SQL = "Select Num_Reg,New_Username,Counter,Max_Online,Max_Tim,Start_Tim,Today_Tim,Num_New,Configs,SuperAdmin,Fso_Sys,AdminFolder,UploadFolder,BadWords,LockIP,UserLevel,Medal,Gbook,ForumData,Payment,UploadFilters From YSvoid_Configs Where ID=1"
			Set Rs=Exec(SQL,1)
			If Rs.Eof Then
				Rs.Close
				Call Exec("Insert Into YSvoid_Configs(Num_Reg,New_Username,Counter,Max_Online,Max_Tim,Start_Tim,Today_Tim,Num_New,Configs,SuperAdmin,Fso_Sys,AdminFolder,UploadFolder,BadWords,LockIP,UserLevel,Medal,Gbook,ForumData,Payment,UploadFilters)",0)
				Set Rs=Exec(SQL,1)
			End If
			Value=Rs.GetRows(1)
			Rs.Close
		End If
		YSvoid_Info = Value
		total_counter=YSvoid_Info(2,0)
		'--- 每日刷新一次统计信息
		If Cstr(time_type(YSvoid_Info(6,0),4))<>Cstr(time_type(now_time,4)) Then
			Call Day_Initialize()
			total_counter=total_counter+1
		End If
		Web_Dir = InstallDir
		Fso_Sys_Var= YSvoid_Info(10,0)
		SuperAdmins = YSvoid_Info(9,0)
		AdminFolder = YSvoid_Info(11,0)
		If Right(AdminFolder,1)<>"/" Then AdminFolder = AdminFolder & "/"
		Web_Upload = YSvoid_Info(12,0)
		Web_BadWords = YSvoid_Info(13,0)
		Web_LockIP = YSvoid_Info(14,0)
		Web_UserLevel = Split(YSvoid_Info(15,0),"|||")
		Web_Medal=Split(YSvoid_Info(16,0),"|||")
		Web_Gbook=YSvoid_Info(17,0)
		Web_Payment=YSvoid_Info(19,0)
		Web_ForumData=YSvoid_Info(18,0)
		Web_UploadFilters=YSvoid_Info(20,0)
		Tmpstr = YSvoid_Info(8,0)
		Tmpstr = Split(Tmpstr,"|||")
		Web_Common = Split(Tmpstr(0),",")
		Web_AdminCodes = Tmpstr(1)
		Web_Number = Split((Tmpstr(2)),",")
		Web_UploadType = Split((Tmpstr(3)),",")
		Web_MailServer = Tmpstr(4)
		Web_RegBadWords = Tmpstr(5)
		Web_RegUserName = Tmpstr(6)
		Web_CopyRight = Tmpstr(7)
		Web_Closer = Tmpstr(8)
		Erase TmpStr
  	End Sub


	'--- 更新总设置表部分缓存数组，入口：更新内容、数组位置
	Public Sub ReloadInfo(MyValue,N)
		YSvoid_Info(N,0) = MyValue
		Name = "YSvoid_info"
		value = YSvoid_Info
		YSvoid_Info = value
	End Sub

	'--- 每日简单初始化
	Private Sub Day_Initialize()
		Dim dTim
		dTim = Time_Type(Now_Time,4)
		Call Exec("Update YSvoid_Configs Set Num_New=0,Today_Tim='"&dTim&"'",0)
		Call Exec("Update YSvoid_BbsForum Set Forum_New_Num=0",0)
		Call Exec("Update YSvoid_UserData Set abate=1 Where estate=1 And otim<"&dTim,0)
		ReloadInfo 0,7
		ReloadInfo dTim,6
	End Sub

	'--- 取缓存总数
	Public Property Get Cache_Num
		Cache_Num=Application.Contents.count
	End Property

	'--- 取用户组总数
	Public Property Get UserGroupLen
		UserGroupLen=0
		Name="user_group"
		If Cache_Chk() Then UserGroup_To_Cache()
		If IsArray(Value) Then
			UserGroupLen=Ubound(Value,2)
		End If
	End Property

	'--- 浏览器前一页URL，返回Encode
	Public Property Get BrowseUrl
		BrowseUrl=Trim(Request.ServerVariables("url"))&"?"&Trim(Request.ServerVariables("QUERY_STRING"))
		BrowseUrl=Server.UrlEncode(BrowseUrl)
	End Property
	Public Property Get WebBrowseUrl
		WebBrowseUrl=Trim(Request.ServerVariables("http_referer"))
	End Property
	Public Property Get CurrentUrl
		CurrentUrl=Trim(Request.ServerVariables("http_referer"))&"?"&Trim(Request.ServerVariables("QUERY_STRING"))
	End Property

	'判断用户是否在线
	Public Function IsONline(vUserName,vAction)
		IsONline=False
		If Trim(vUserName)="" Then Exit Function
		If IsArray(session_get(vUserName)) And vAction=1 Then
				IsONline=True:Exit Function
		End If
		Dim oRs
		Set oRs=Exec("Select Count(*) From YSvoid_UserLogin Where l_username='"&vUserName&"'",1)
		If oRs(0)<>0 Then IsONline=True
		oRs.Close:Set oRs = Nothing
	End Function

	'更新用户资料缓存(缓存用户名,是否需要添加)[0=不添加,只作清理,1=需要添加]
	Public Sub NeedUpdateList(vUserName,vAct)
		Dim Tmpstr,TmpUsername
		Name="NeedToUpdate"
		If Cache_Chk() Then Value=""
		Tmpstr=Value
		TmpUsername=","&vUserName&","
		Tmpstr=Replace(Tmpstr,TmpUsername,",")
		Tmpstr=Replace(Tmpstr,",,",",")
		IF vAct=1 Then
			If IsONline(vUserName,0) Then
				If Tmpstr="" Then
					Tmpstr=TmpUsername
				Else
					Tmpstr=Tmpstr&TmpUsername
				End If
			End If
		End If
		Tmpstr=Replace(Tmpstr,",,",",")
		Value=Tmpstr
	End Sub

	'--- 论坛ID ---
	Public Property Get Forum_ID
		Forum_ID=Code_id("fid")
	End Property

	'--- 页面执行时间 ---
	Public Property Get Exec_Tim
		Exec_Tim=FormatNumber((Timer()-Timer_Start),3)
		If Left(Exec_Tim,1)="." Then Exec_Tim="0"&Exec_Tim
	End Property

	'//* 模板操作部分开始 *//
	'-----------模板列表---------------
	Private Sub MainSkinList()
		Dim oSQL, oRs
		oSQL = "Select TemplateID,TemplateName From [YSvoid_Template] Where ChannelID=0"
		Set oRs = Exec(oSQL,1)
		Value = oRs.GetRows(-1)
		oRs.Close:Set oRs=Nothing
	End Sub

	'-----------取默认模板 ID-----------
	Private Sub TemplateID_Load()
		Dim oSQL, oRs
		oSQL = "Select Top 1 TemplateID From [YSvoid_Template] Where IsDefault=True And ChannelID=0"
		Set oRs = Exec(oSQL,1)
		If Not oRs.Eof Then
			Value = Int(oRs(0))
			oRs.Close:Set oRs=Nothing
			Exit Sub
		End If
		oRs.Close
		oSQL = "Select Top 1 TemplateID From [YSvoid_Template] Where ChannelID=0"
		Set oRs = Exec(oSQL,1)
		If Not oRs.Eof Then
			Value = Int(oRs(0))
			oRs.Close:Set oRs=Nothing
			Call Exec("Update [YSvoid_Template] Set IsDefault=True Where TemplateID="&Value&" And ChannelID=0",1)
			Exit Sub
		Else
			oRs.Close:Set oRs=Nothing
			ClsErr("模板数据是空的，请先添加。")
		End If
	End Sub

	'-----------取模板数据---------
	Private Sub TemplateCache(ByVal tChannelID, ByVal tSelectID)
		Dim oSQL, oRs, nChannelID, nSQLAdd, nSelSQL
		nSQLAdd = " And IsDefault=True"
		nSelSQL = "TemplateConfig,TemplateWhole,TemplateOther,TemplateStr,SkinCss"
		If tChannelID = 0 Then
			nSQLAdd = " And TemplateID="&tSelectID
			nSelSQL = "TemplateName,TemplateIndex,TemplateWhole,TemplateOther,TemplateStr,SkinDir"
		End If
		oSQL = "Select Top 1 "&nSelSQL&" From [YSvoid_Template] Where ChannelID="&tChannelID&nSQLAdd&" Order By TemplateID Desc"
		Set oRs = Exec(oSQL,1)
		If Not oRs.Eof Then
			Value = oRs.GetRows()
			oRs.Close:Set oRs=Nothing
			Exit Sub
		End If
		oRs.Close
		oSQL = "Select Top 1 TemplateID From [YSvoid_Template] Where ChannelID="&tChannelID&" Order By TemplateID Desc"
		Set oRs = Exec(oSQL,1)
		If Not oRs.Eof Then
			Call Exec("Update [YSvoid_Template] Set IsDefault=True Where TemplateID="&oRs(0)&" And ChannelID="&tChannelID,1)
			Value = oRs.GetRows()
			oRs.Close:Set oRs=Nothing
			Exit Sub
		Else
			oRs.Close:Set oRs=Nothing
			'处理错误
			If tChannelID = 0 Then
				Response.Redirect "Skins.asp?Action="&iMainSkinID
			Else
				ClsErr("频道模板数据是空的，请先添加。")
			End If
		End If
	End Sub

	'--------分配模板数组------------
	Private Sub MainMod_Load()
		If IsDeBug = 1 Then On Error Resume Next'容错代码。
		Dim AllStr, Tmpstr, iSkinID, iCookiesSid
		'-----检查默认模板ID
		Name = "SkinID"
		If Cache_Chk() Then TemplateID_Load()
		iMainSkinID = Value
		iCookiesSid = Cookies_Get("skin_id")
		If Not Is_Int(iCookiesSid) Then
			iSkinID=iMainSkinID
		Else
			iSkinID=iCookiesSid
		End If
		iSkinID=CLng(iSkinID)
		'-----取模板数组
		Name="MainStyle"&iSkinID
		If Cache_Chk() Then Call TemplateCache(0,iSkinID)
		AllStr = Value
		If Not IsArray(AllStr) Then
			ClsErr("频道模板数据导常！请联系管理员！")
			Exit Sub
		End If
		If Ubound(AllStr,1)<>5 Then
			ClsErr("频道模板数据导常！请联系管理员！")
			Exit Sub
		End If
		'-----取CSS路径
		iMainCssPath = YSvoid.Web_Dir&"Skin/Skin_"&iSkinID&".css"
		iMainCssPath = "<link type=""text/css"" href="""&iMainCssPath&""" rel=stylesheet>"
		'-----取模板名称
		iMainName = AllStr(0,0)
		'-----取首页模板
		iMainIndex = AllStr(1,0)
		'-----取模板框架
		iMainWhole = Split(AllStr(2,0),"|||")
		'-----取其他模板
		Tmpstr = Split(AllStr(3,0),"@@@")(0)
		iMainOther = Split(Tmpstr,"|||")
		Tmpstr = Split(AllStr(3,0),"@@@")(1)
		iMainOtherName = Split(Tmpstr,"|||")
		'-----取字符串
		Tmpstr = Split(AllStr(4,0),"@@@")(0)
		iMainStrer = Split(Tmpstr,"|||")
		Tmpstr = Split(AllStr(4,0),"@@@")(1)
		iMainStrerName = Split(Tmpstr,"|||")
		'-----取模板页面宽度
		iMainWidth = iMainWhole(12)
		'-----取图片路径
		iMainPicPath = Split(AllStr(5,0),"|||")(0)
		Erase AllStr
		web_dir_skin = YSvoid.Web_Dir&"Skin/"&iMainPicPath&"/"
	End Sub

	'--- 判断栏目模板是否载入并赋值
	Public Function Channel_TemplateLoad(ByVal tChannelID)
		Dim iPageCss
		iPageCss = ""
		Channel_TemplateLoad=False
		If Is_Null(tChannelID) = "" Or tChannelID = 0 Then Exit Function
		Name = "ChannelStyle_"&tChannelID
		If Cache_Chk() Then Call TemplateCache(tChannelID,0)
		Dim AllStr, Tmpstr
		AllStr = Value
		If Not IsArray(AllStr) Then
			Exit Function
		End If
		If Ubound(AllStr,1)<>4 Then
			Erase AllStr
			Exit Function
		End If
		'-----取模板框架，否则使用全局框架
		If AllStr(0,0) = True Then
			iMainWhole = Split(AllStr(1,0),"|||")
			'-----取模板页面宽度
			iMainWidth = iMainWhole(12)
		End If
		'-----取其他模板
		Tmpstr = Split(AllStr(2,0),"@@@")(0)
		iPageOther = Split(Tmpstr,"|||")
		Tmpstr = Split(AllStr(2,0),"@@@")(1)
		iPageOtherName = Split(Tmpstr,"|||")
		'-----取字符串，可有于多语言设置
		Tmpstr = Split(AllStr(3,0),"@@@")(0)
		iPageStrer = Split(Tmpstr,"|||")
		Tmpstr = Split(AllStr(3,0),"@@@")(1)
		iPageStrerName = Split(Tmpstr,"|||")
		iPageCss = AllStr(4,0)
		If iPageCss <> "" Then
			iMainCssPath = iMainCssPath & Vbcrlf & "<style>"
			iMainCssPath = iMainCssPath & Vbcrlf & "<!-- 频道 CSS "
			iMainCssPath = iMainCssPath & Vbcrlf & iPageCss
			iMainCssPath = iMainCssPath & Vbcrlf & "-->"
			iMainCssPath = iMainCssPath & Vbcrlf & "</style>"
		End If
		Erase AllStr
		Channel_TemplateLoad = True
	End Function

	'--- 模板中常用配置替换 ---
	Public Function ReHtml(HtmlStr,HtmlType)
		Dim Rii
		ReHtml = Replace(Cstr(HtmlStr),"{$lefter}",iMainWhole(2))
		ReHtml = Replace(ReHtml,"{$righter}",iMainWhole(3))
		ReHtml = Replace(ReHtml,"{$skin_dir}",web_dir_skin)
		ReHtml = Replace(ReHtml,"{$skin_name}",iMainName)
		ReHtml = Replace(ReHtml,"{$skin_css}",iMainCssPath)
		ReHtml = Replace(ReHtml,"{$web_dir}",Web_Dir)
		ReHtml = Replace(ReHtml,"{$upload_dir}",Web_Upload)
		ReHtml = Replace(ReHtml,"{$download_dir}",Web_Dir&Web_Upload)
		ReHtml = Replace(ReHtml,"{$web_width}",iMainWidth)
		ReHtml = Replace(ReHtml,"{$web_url}",Web_Url)
		ReHtml = Replace(ReHtml,"{$web_name}",Web_Name)
		ReHtml = Replace(ReHtml,"{$edition}",pro_edition)
		ReHtml = Replace(ReHtml,"{$web_email}",Web_Common(6))
		ReHtml = Replace(ReHtml,"{$web_tit}",tit)
		ReHtml = Replace(ReHtml,"{$web_unit}",Web_Common(8))
		ReHtml = Replace(ReHtml,"{$web_br}",""&vbcrlf&"")
		ReHtml = Replace(ReHtml,"{$pro_edtion}",pro_edition)
		ReHtml = Replace(ReHtml,"{$web_copyright}",Web_CopyRight)
		ReHtml = Replace(ReHtml,"{$admin_dir}",AdminFolder)
		If HtmlType=0 Then
			For Rii=0 To UBound(iMainStrer)-1
				ReHtml=Replace(ReHtml,"{$"&iMainStrerName(Rii)&"}",iMainStrer(Rii))
			Next
		End If
	End Function

	Public Function ModGet(Mname)
		ModGet=""
		Select Case Lcase(Mname)
		Case "lefter"
			ModGet=iMainOther(2)
			Exit Function
		Case "righter"
			ModGet=iMainOther(3)
			Exit Function
		End Select
		ModGet=HtmlStream(Mname,iMainStrerName,iMainStrer)
		'ModGet=ReHtml(ModGet,1)
	End Function

	'--- 取全局模板(数组方式) ---
	Public Sub HtmlWhole(Mnum)
		PutHtml=iMainWhole(Mnum)
	End Sub

	'--- 取全局模板(命名方式) ---
	Public Sub HtmlMain(Mname)
		PutHtml=HtmlStream(Mname,iMainOtherName,iMainOther)
	End Sub

	'--- 取分页模板(数组方式) ---
	Public Sub HtmlPageNum(Mnum)
		PutHtml=iPageOther(Mnum)
	End Sub

	'--- 取分页模板(命名方式) ---
	Public Sub HtmlPage(Mname)
		PutHtml=HtmlStream(Mname,iPageOtherName,iPageOther)
	End Sub

	'--- 命名方式读取HTML流数据
	Private Function HtmlStream(ByVal Mname,ByVal sName,ByVal sValue)
		HtmlStream=""
		If Mname="" Then Exit Function
		If Not IsArray(sName) Or Not IsArray(sValue) Then Exit Function
		If Ubound(sName)<>Ubound(sValue) Then Exit Function
		Dim Sii
		For Sii=0 To UBound(sName)-1
			If LCase(sName(Sii))=LCase(Mname) Then
				HtmlStream=sValue(Sii)
				Exit For
			End If
		Next
	End Function

	'--- 模板标签替换 ---
	Public Sub HtmlRcod(Mod_Str,View_Str)
		If IsDeBug=0 Then On Error Resume Next
		'Mod_Str=LCase(Mod_Str)
		PutHtml=Replace(PutHtml,"{$"&Mod_Str&"}",View_Str)
	End Sub

	'--- 模板打印 ---
	Public Sub HtmlView(Knum)
		Response.Write ReHtml(PutHtml,0) & VbCrlf
		'Response.Write ReHtml(PutHtml,0)
		If Knum=1 Then Response.Write ModGet("ukong")
	End Sub

	'--- 模板赋值 ---
	Public Function HtmlGet(Knum)
		HtmlGet=ReHtml(PutHtml,0) & VbCrlf
		'HtmlGet=ReHtml(PutHtml,0)
		If Knum=1 Then HtmlGet=HtmlGet & ModGet("ukong")
	End Function
	'//* 模板操作结束 *//

	'//* 网站、菜单频道信息处理开始 *//
	'//* mType:( 0=前台；1=后台 ) *//
	Private Sub MenuToCache(ByVal mType)
		Dim oRs, oSQL
		If mType = 0 Then
			oSQL="Select * From [YSvoid_Channel] Where ChannelHidden=False Order By ChannelOrder,ChannelID"
		Else
			oSQL="Select * From [YSvoid_Menu] Order By ChannelID,MenuOrder,MenuID"
		End If
		Set oRs=Exec(oSQL,1)
		If Not oRs.Eof Then
			Value=oRs.GetRows(-1)
		Else
			Value=""
		End If
		oRs.Close:Set oRs=Nothing
	End Sub

	'--- 取频道常用信息
	Public Property Get Channel_Menu
		Channel_Menu=Cache_Get("web_menu")
	End Property
	'//* 网站、菜单频道信息处理结束 *//

	'//* 频道信息处理 *//
	'--- 频道初始化操作
	Public Sub Channel_Initialize(ByVal nChannelID)
		Dim ChannelDim, Cii, iUpFileType
		If nChannelID =< 0 Then Exit Sub
		'--- 载入前台频道设置
		ChannelDim = Channel_Menu
		'--- 遍列菜单设置数据并进行赋值
		For Cii=0 To Ubound(ChannelDim,2)
			If Int(nChannelID)=Int(ChannelDim(0,Cii)) Then
				ChannelName			= Code_Html(ChannelDim(1,Cii),1,0)
				ChannelDir			= Replace(Replace(ChannelDim(12,Cii),"/",""),"\","")
				ChannelTit			= Code_Html(ChannelDim(13,Cii),1,0)
				ChannelTable		= ChannelDim(14,Cii)
				ChannelUnit			= Code_Html(ChannelDim(15,Cii),1,0)
				ChannelListNum		= Split(ChannelDim(16,Cii),"@@@")
				ChannelPosition		= Split(ChannelDim(20,Cii),"@@@")
				ChannelSetup		= Cstr(ChannelDim(18,Cii))
				ClassDepth			= ChannelDim(19,Cii)
				'ChannelPutType		= Int(ChannelDim(21,Cii))
				ChannelUpType		= Int(ChannelDim(22,Cii))
				ChannelUpMax		= Int(ChannelDim(23,Cii))
				ChannelPicWidth		= Int(ChannelDim(24,Cii))
				ChannelPicHeight	= Int(ChannelDim(25,Cii))
				ChannelUpFileType	= Cstr(ChannelDim(26,Cii))
				ChannelCuteNum		= Split(ChannelDim(27,Cii),"@@@")
				ChannelLeftCuteNum	= Int(ChannelDim(28,Cii))
				ChannelCount		= ChannelDim(29,Cii)
				ChannelHits			= ChannelDim(30,Cii)
				ChannelOption		= ChannelDim(31,Cii)
				ChannelUploadDir    = ChannelDim(32,Cii)
				ChannelReward		= Split(ChannelDim(33,Cii),"@@@")
				ChannelDates		= Split(ChannelDim(34,Cii),"@@@")
				ChannelRss          = ChannelDim(35,Cii)
                ChannelPower        = ChannelDim(36,Cii)
				ChannelPHidden      = ChannelDim(37,Cii)
				ChannelTrue			= True
				Exit For
			End If
		Next
		Erase ChannelDim
		If Not ChannelTrue Then Exit Sub
		If ClassDepth = 0 Then ClassDepth = 99999999
		iUpFileType = Split(ChannelUpFileType,"@@@")
		If Not IsArray(iUpFileType) Then
			ChannelUpFileType = Split("@@@@@@@@@@@@","@@@")
		Else
			ChannelUpFileType = iUpFileType
			Erase iUpFileType
		End If
		If Get_ChannelSetup(ChannelSetup,2) = 1 Then ReviewTrue = True
		If Get_ChannelSetup(ChannelSetup,3) = 1 Then SpecialTrue = True
		If Get_ChannelSetup(ChannelSetup,13) = 1 Then CastTrue = True
		If Get_ChannelSetup(ChannelSetup,4) = 1 Then KeyWordTrue = True
		If Get_ChannelSetup(ChannelSetup,5) = 1 Then ErrTrue = True
		If Get_ChannelSetup(ChannelSetup,8) = 1 Then IsHealth = True
	End Sub

	'--- 获取某个频道的某个数组值
	Public Function Get_ChannelRes(ByVal nChannelID, ByVal iNum)
		Dim Rii, sChannel
		Get_ChannelRes = ""
		If nChannelID = 0 Or Not Is_Int(iNum) Then Exit Function
		sChannel = Channel_Menu
		If Int(iNum) > Int(Ubound(sChannel,1)) Then
			Exit Function
			Erase sChannel
		End If
		For Rii = 0 To Ubound(sChannel,2)
			If Int(nChannelID) = sChannel(0,Rii) Then
				Get_ChannelRes = sChannel(iNum,Rii)
				Exit For
			End If
		Next
		Erase sChannel
	End Function

	'--- 读取频道参数值
	'--- 0-发布权、1-分类、2-评论、3-专题、4-关键字、5-报错
	'--- 6-RSS、7-排版样式、8-过滤字符、9-前台编辑器、10-后台编辑器、13-公告
	Public Function Get_ChannelSetup(ByVal cSetup,ByVal sNum)
		Get_ChannelSetup = 0
		sNum = sNum + 1
		If Len(cSetup) <> 20 Or sNum > 20 Then Exit Function
		Get_ChannelSetup = Int(Mid(cSetup,sNum,1))
	End Function

	'//* 频道专题、关键字、公告处理 *//
	Public Sub Special_KeyWord_Cast_ToCache(ByVal oChannelID,ByVal oType)
		Dim oRs, oSQL
		If oChannelID = 0 Then Exit Sub
		oSQL="Select * From [YSvoid_Special] Where ChannelID="&oChannelID&" Order By ID"
		If oType = 1 Then
			oSQL="Select * From [YSvoid_KeyWord] Where ChannelID="&oChannelID&" Order By ID"
		End If
		If oType = 2 Then
			oSQL="Select * From [YSvoid_Cast] Where ChannelID="&oChannelID&" Order By ID"
		End If
		Set oRs=Exec(oSQL,1)
		If Not oRs.Eof Then
			Value=oRs.GetRows()
		Else
			Value=""
		End If
		oRs.Close:Set oRs=Nothing
	End Sub

	'//* 模板缓存操作部分开始 *//
	'--- 转换缓存名称小写
	Public Property Let Name(ByVal vNewValue)
		LocalCacheName=LCase(vNewValue)
	End Property

	'--- 缓存赋值，数组方式
	Public Property Let Value(ByVal vNewValue)
		If LocalCacheName<>"" Then
			ReDim Cache_Data(2)
			Cache_Data(0)=vNewValue
			Cache_Data(1)=Now()
			Application.Lock
			Application(Web_Cookies & "_" & LocalCacheName) = Cache_Data
			Application.unLock
		Else
			ClsErr("缓存名称未设置 或 与已有的重复！请重新设置！")
		End If
	End Property

	'--- 缓存取值，数组0
	Public Property Get Value()
		If LocalCacheName<>"" Then
			Cache_Data=Application(Web_Cookies & "_" & LocalCacheName)
			If IsArray(Cache_Data) Then
				Value=Cache_Data(0)
			Else
				ClsErr("缓存 ("&LocalCacheName&") 没有内容 ！")
			End If
		Else
			ClsErr("缓存名称未设置 或 与已有的重复！请重新设置！")
		End If
	End Property

	'//* 打印类错误页面 *//
	Private Sub ClsErr(ClsStr)
		Response.Clear
		Response.Write ClsStr
		Response.End
	End Sub


	'//* 数据库操作部分开始 *//
	Public Function Exec(Esql,Etype)
		If IsConn=False Then Call Conn_Open()
		If IsDeBug=1 Then On Error Resume Next
		Select Case Etype
		Case 0
			Conn.Execute(Esql)
		Case 1
			Set Exec=Conn.Execute(Esql)
		End Select
		If IsView=1 Then Response.Write "错误语句："&esql&"<br>"
		'If IsView=1 Then Call ClsErr("错误语句："&esql&"<br>")
		If Err And IsDeBug=1 Then
			Err.Clear
			Response.Write "错误语句："&esql &"<br>"
			Response.End
		End If
		Num_Rs=Num_Rs+1
	End Function

	Public Function Exe_Conn(Ers,Esql,Etype)
		If IsConn=False Then Call Conn_Open()
		If IsDeBug=1 Then On Error Resume Next
		Ers.Open Esql,Conn,1,Etype
		If IsView=1 Then Response.Write "错误语句："&esql&"<br>"
		If Err And IsDeBug=1 Then
			Err.Clear
			Response.Write "错误语句："&esql&"<br>"
			Response.End
		End If
		Num_Rs=Num_Rs+1
	End Function

	Private Sub Conn_Open()
		Set Conn=Server.CreateObject("Adodb.Connection")
		Set Rs=Server.CreateObject("Adodb.Recordset")
		Conn.Open Connstr
		IsConn=True
	End Sub

    '函数:SqlServer(97-2000) to Access(97-2000)
    '参数:Sql,数据库类型(ACCESS,SQLSERVER)
    '说明:
    Private Function Sql_To_Access(ByVal vNewSql)
        Dim regEx, Matches, Match, vSql
        vSql=vNewSql
        '创建正则对象
        Set regEx = New RegExp
        regEx.IgnoreCase = True
        regEx.Global = True
        regEx.MultiLine = True
        '转:GetDate()
        regEx.Pattern = "(?=[^']?)GETDATE\(\)(?=[^']?)"
        vSql = regEx.Replace(vSql,"NOW()")
        '转:UPPER()
        regEx.Pattern = "(?=[^']?)UPPER\([\s]?(.+?)[\s]?\)(?=[^']?)"
        vSql = regEx.Replace(vSql,"UCASE($1)")
        '转:日期表示方式
        '说明:时间格式必须是2004-23-23 11:11:10 标准格式
        regEx.Pattern = "'([\d]{4,4}\-[\d]{1,2}\-[\d]{1,2}(?:[\s][\d]{1,2}:[\d]{1,2}:[\d]{1,2})?)'"
        vSql = regEx.Replace(vSql,"#$1#")
        regEx.Pattern = "DATEDIFF\([\s]?(mi|second|minute|hour|d|month|year)[\s]?\,[\s]?(.+?)[\s]?\,[\s]?(.+?)([\s]?\)[\s]?)"
        Set Matches = regEx.ExeCute(vSql)
        Dim temStr
        For Each Match In Matches
            temStr = "DATEDIFF("
            Select Case lcase(Match.SubMatches(0))
                Case "mi" :
                    temStr = temStr & "'n'"
                Case "second" :
                    temStr = temStr & "'s'"
                Case "minute" :
                    temStr = temStr & "'n'"
                Case "hour" :
                    temStr = temStr & "'h'"
                Case "d" :
                    temStr = temStr & "'d'"
                Case "month" :
                    temStr = temStr & "'m'"
                Case "year" :
                    temStr = temStr & "'y'"
            End Select
            temStr = temStr & "," & Match.SubMatches(1) & "," &  Match.SubMatches(2) & Match.SubMatches(3)
            vSql = Replace(vSql,Match.Value,temStr,1,1)
        Next
        '转:Insert函数
        regEx.Pattern = "CHARINDEX\([\s]?'(.+?)'[\s]?,[\s]?'(.+?)'[\s]?\)[\s]?"
        vSql = regEx.Replace(vSql,"INSTR('$2','$1')")
        Set regEx = Nothing
        Sql_To_Access = vSql
    End Function

  function popedom_p(popedom_var,popedom_n)
    if cint(popedom_n)>len(popedom_var) or len(popedom_var)<>120 then
      popedom_p=0
      exit function
    End If
    popedom_p=int(mid(popedom_var,popedom_n,1))
  End Function

  '-------------------------------------地址处理--------------------------------------
  function url_true(puu,pus)
    dim pupload,pu
    pupload=puu
    pu=pus
    if instr(1,pu,"://")<>0 then
      url_true=pu
      'url_true=Web_Dir&pu
    else
      if pupload="" then pupload=web_upload
      if right(pupload,1)<>"/" then pupload=pupload&"/"
      url_true=Web_Dir&pupload&pu
    End If
  End Function

  function own_http()
    dim ssort
    ssort=trim(request.servervariables("server_port"))
    own_http="http://"&trim(request.servervariables("server_name"))
    if ssort<>"80" then own_http=own_http&":"&ssort
    if right(own_http,1)<>"/" then own_http=own_http&"/"
  End Function

  sub own_dir()
    dim path_info,now_dir,ndd
    path_info=request.servervariables("path_info")
    now_dir=left(path_info,instrrev(path_info,"/"))
    web_now_dir=now_dir
    ndd=own_http()
	Web_RealURL=left(ndd,len(ndd)-1)	
    web_urls=Web_RealURL&now_dir
  End Sub

  '-----------------------------------时间格式处理------------------------------------
  function time_type(tvar,tt)
    dim ttt,d_year,d_month,d_day,d_hour,d_minute,d_second
    ttt=tvar
    if ttt="" or isnull(ttt) then ttt=now_time
    if not(isdate(ttt)) then
      time_type=""
      exit function
    End If
    d_year=year(ttt)
    d_month=month(ttt)
    if len(d_month)<2 then d_month="0"&d_month
    d_day=day(ttt)
    if len(d_day)<2 then d_day="0"&d_day
    d_hour=hour(ttt)
    if len(d_hour)<2 then d_hour="0"&d_hour
    d_minute=minute(ttt)
    if len(d_minute)<2 then d_minute="0"&d_minute
    d_second=second(ttt)
    if len(d_second)<2 then d_second="0"&d_second
    select case tt
    case 1	'2000-10-10 23:45:45
      time_type=d_year&"-"&d_month&"-"&d_day&" "&d_hour&":"&d_minute&":"&d_second
    case 11	'20001010234545
      time_type=d_year&d_month&d_day&d_hour&d_minute&d_second
    case 2	'年(4)-月-日 时:分:秒
      time_type=d_year&"年"&d_month&"月"&d_day&"日 "&d_hour&"时"&d_minute&"分"&d_second&"秒"
    case 3	'10-10 23:45
      time_type=d_month&"-"&d_day&" "&d_hour&":"&d_minute
    case 4	'2003-10-10
      time_type=d_year&"-"&d_month&"-"&d_day
    case 5	'2003年10月10日
      time_type=d_year&"年"&d_month&"月"&d_day&"日"
    case 6	'10-10
      time_type=d_month&"-"&d_day
    case 7	'03-10-10
      time_type=right(d_year,len(d_year)-2)&"-"&d_month&"-"&d_day
    case else
      time_type=ttt
    end select
  End Function

	'//* 字符串处理 *//
'-- 判断字符串是否为空
	Public Function Is_Null(ByVal sVal)
		Is_Null=Trim(sVal)
		If Is_Null="" Or IsNull(Is_Null) Then Is_Null=""
	End Function

	'-- 判断是否为整数类型，返回布尔值
	Public Function Is_Int(ByVal sVal)
		Is_Int=True
		If Is_Null(sVal)="" Or Not(IsNumeric(sVal)) Or Instr(sVal,".")>0 Then Is_Int=False
	End Function

	'-- 判断是否为数字类型，返回布尔值
	Public Function Is_Num(ByVal sVal)
		Is_Num=True
		If Is_Null(sVal)="" Or Not(IsNumeric(sVal)) Then Is_Num=False
	End Function

	'-- E-Mail判断
	Public Function Is_Email(ByVal sEmail,ByVal sEn)
		Dim names,name,ei,c
		Is_Email=False
		names=Split(sEmail,"@")
		If Ubound(names)<>1 Or sEmail="" Then
			Exit Function
		End If
		For Each name In names
			If Len(name) <= 0 Then
				Exit Function
			End If
			For ei=1 To Len(name)
				c=LCase(Mid(name,ei,1))
				If Instr("abcdefghijklmnopqrstuvwxyz-_.",c)<= 0 And Not(IsNumeric(c)) Then
					Exit Function
				End If
			Next
			If Left(name,1)="." Or Right(name,1)="." Then
				Exit Function
			End If
		Next
		If Instr(names(1),".")<=0 Or Len(sEmail)>sEn Then
			Exit Function
		End If
		Is_Email=True
	End Function

	'-- 字符处理
	Public Function Code_Var(ByVal sVal)
		Dim Strer
		Strer=Is_Null(sVal)
		If Strer="" Then
			Code_Var=""
			Exit Function
		End If
		Strer=Replace(Strer,"'","""")
		Code_Var=Strer
	End Function
	
	
  Public Function Code_Encrypt(ByVal sVal)
		Dim Strer
		Strer=Is_Null(sVal)		
		If Strer="" Then
			Code_Encrypt=""
			Exit Function
		End If		
		Code_Encrypt=Strer
  End Function

	'-- 处理 Form 值
	Public Function Code_Form(ByVal sVal)
		Dim Strer
		Strer=Trim(Request.Form(sVal))
		Strer=Replace(Strer,"'","""")
		Code_Form=Strer
	End Function
	
	Public Function Code_Sel(ByVal sVal, ByVal sType)
		Dim Strer, TmpStr, Tmp, TmpDim, m
		For Each Strer In Request.Form(sVal)
			TmpStr = Replace(Replace(Strer,",","")," ","")
			If sType = 1 Then
				If Not Is_Int(TmpStr) Then TmpStr = ""
			End If
			If TmpStr<>"" Then
				If Tmp = "" Then
					Tmp = TmpStr
				Else
					Tmp = Tmp & "," & TmpStr
				End If
			End If
		Next
		Code_Sel = Tmp
	End Function

	'-- 处理 Query 值
	Public Function Code_Query(ByVal sVal)
		Dim Strer
		Strer=Trim(Request.Querystring(sVal))
		Strer=Replace(Strer,"'","")
		Strer=Replace(Strer,"""","")
		Code_Query=Strer
	End Function

	Public Function Code_Remark(ByVal sVal)
		Dim Strer
		Strer=Request.Form(sVal)
 		Strer=Replace(Strer,"","&quot;")		'双引号
        Strer=Replace(strer,Chr(33),"&nbsp;")		'空格
		Code_Remark=Strer
	End Function

	'-- 处理 ID 传入值
	Public Function Code_Id(ByVal sVal)
		Dim Strer
		Strer=Trim(Request.Querystring(sVal))
		If Not(Is_Int(Strer)) Then Strer=0
		Code_Id=Clng(Strer)
	End Function

	'-- 处理整数值
	Public Function Code_Int(ByVal Strers,ByVal Sint)
		Dim Strer
		Strer=Code_Form(Strers)
		If Not(Is_Int(Strer)) Then
			Code_Int=Sint
			Exit Function
		End If
		Code_Int=Int(Strer)
	End Function

	'取各种返回值
	'Strers:字符串
	'At    :类型
	'Acut  :截短字符串数
	Public Function Code_Admin(ByVal Strers,ByVal At,ByVal Acut)
		Dim Strer
		Strer=Trim(Strers)
		Select Case Int(At)
		Case 1
			Strer=Trim(Request.Form(Strer))
		Case 2
			Strer=Trim(Request.QueryString(Strer))
		End Select
		If Is_Null(Strer)="" Then
			Code_Admin=""
			Exit Function
		End If
		Strer=Replace(Strer,"'","""")
		If Int(Acut)>0 Then Strer=Left(Strer,Acut)
		Code_Admin=Strer
	End Function

	Public Function Code_Word(ByVal Strers)
		Dim Strer
		Strer=Trim(Strers)
		If Is_Null(Strer)="" Then
			Code_Word=""
			Exit Function
		End If
		Strer=Replace(Strer,Chr(39),"&#39;")		'单引号
		Strer=Replace(Strer,"<","&lt;")
		Strer=Replace(Strer,">","&gt;")
		Code_Word=Strer
	End Function

	Public Function Code_Js(ByVal Strers,ByVal Tt)
		Dim Strer
		Strer=Trim(Strers)
		If Is_Null(Strer)="" Then
			Code_Js=""
			Exit Function
		End If
		Strer=Code_Health(Strer)		'过滤不健康信息
		Strer=Replace(Strer,"\","\\")
		Strer=Replace(Strer,Chr(39),"&#39;")		'单引号
		'Strer=Replace(Strer,Chr(39),"\'")
		Strer=Replace(Strer,Chr(34),"&quot;")		'双引号
		'Strer=Replace(Strer,Chr(34),"\""")
		Select Case Tt
		Case 0
			'Strer=Replace(Strer,VbCrlf,"")			'回车
			Strer=Replace(Strer,Chr(10),"")
			Strer=Replace(Strer,Chr(13),"")
		Case 1
			'Strer=Replace(Strer,VbCrlf,"\n")			'回车
			Strer=Replace(Strer,Chr(10),"\n")
			Strer=Replace(Strer,Chr(13),"")
		End Select
		If Right(Strer,1)="\" Then
			Strer=Strer&"n"
		End If
		Code_Js=Strer
	End Function

	Public Function Code_Html(ByVal Strers,ByVal Chtype,ByVal CuteNum)
		Dim Strer
		Strer=Trim(Strers)
		If Is_Null(Strer)="" Then
			Code_Html=""
			Exit Function
		End If
		Strer=Code_Health(Strer)		'过滤不健康信息
		If CuteNum>0 Then Strer=Cuted(Strer,CuteNum)
		Strer=Replace(Strer,"<","&lt;")
		Strer=Replace(Strer,">","&gt;")
		Strer=Replace(Strer,Chr(39),"&#39;")		'单引号
		Strer=Replace(strer,Chr(34),"&quot;")		'双引号
		Strer=Replace(strer,Chr(33),"&nbsp;")		'空格
		Select Case Chtype
		Case 1
			Strer=Replace(Strer,Chr(9),"&nbsp;")		'table
			'Strer=Replace(Strer,VbCrlf,"")			'回车
            Strer=Replace(Strer,"{$download_dir}",Web_Dir&Web_Upload)
			Strer=Replace(Strer,Chr(10),"")
			Strer=Replace(Strer,Chr(13),"")
		Case 2
			Strer=Replace(Strer,Chr(9),"&nbsp;　&nbsp;")	'table
			'Strer=Replace(Strer,VbCrlf,"<br \>")		'回车
			Strer=Replace(Strer,Chr(10),"<br \>")
			Strer=Replace(Strer,Chr(13),"")
		End Select
		Code_Html=Strer
	End Function

	'过滤不健康的字符串
	Public Function Code_Health(ByVal Strer)
		Dim i,BadWords,iBadWords,lBadWords,rBadWords
		Code_Health=Strer
		If Not IsHealth Then Exit Function
		BadWords=Split(Web_BadWords,"|")
		For i=0 To Ubound(BadWords)
		  iBadWords=Split(BadWords(i),"=")
		  lBadWords=iBadWords(0)
		  rBadWords=iBadWords(1)
		  Code_Health=Replace(Code_Health,lBadWords,rBadWords)
		Next
		Erase BadWords
	End Function

	    '过滤非法的SQL字符
	Public Function ReplaceBadChar(Byval strChar)
		strChar=replace(replace(strChar," ",""),"'","")
		strChar=replace(replace(strChar,".",""),"<","")
		strChar=replace(replace(strChar,")",""),"(","")
		strChar=replace(replace(strChar,"?",""),"*","")
		strChar=replace(replace(strChar,"/",""),"\","")
		ReplaceBadChar=replace(strChar,Chr(0),"")
	End Function
	
    '能识别中文字的通用字符分割器(取指定长度的字符，标题显示用）OK
    'Types      为输入的待分割字符
    'Num        为分割的长(字节数，中文为２字节),左起第几个
    Public Function Cuted(ByVal Types,ByVal Num)
        If Is_null(Types) = "" Or Types = "" Then
            Cuted = ""
            Exit Function
        Else
            If Str_Length(Types) > Cint(Num) Then
                Types = Str_Left(Types,Cint(Num))
                Cuted = Types & "..."
            Else
                Cuted = Types
            End If
        End If
    End Function

    '能识别中文字的left(),中文为２字节 OK
    'Strers   为原字符串
    'Num     为需要(左)截取的位数
    Public Function Str_Left(ByVal Strers,ByVal Num)
        Dim Temp_Str, Test_Str, Lens, i
        Temp_Str = Len(Strers)
        For i = 1 To Temp_Str
            Test_Str = Mid(Strers, i, 1)
            Str_Left = Str_Left & Test_Str
            If Asc(Test_Str) > 0 Then
                Lens = Lens + 1
            Else
                Lens = Lens + 2
            End If
            If Lens >= Cint(Num) Then
                Exit For
            End If
        Next
    End Function

    '能识别中文字的len(),中文为２ OK
    'Strers    为待检测长度的字符串
    Public Function Str_Length(ByVal Strers)
        On Error Resume Next
        Dim Winnt_Chinese
        Winnt_Chinese = (Len("中国") = 2)
        If Winnt_Chinese = True Then
            Dim LNum, TNum, CNum, i
            LNum = Len(Strers)
            TNum = LNum
            For i = 1 To LNum
                CNum = Asc(Mid(Strers, i, 1)) '返回字符的代码值
                If CNum < 0 Then
                    CNum = CNum + 65536
                End If
                If CNum > 255 Then
                    TNum = TNum + 1
                End If
            Next
            Str_Length = TNum
        Else
            Str_Length = Len(Strers)
        End If
        If Err.Number <> 0 Then
            Err.Clear()
        End If
    End Function

	'显示验证码
	Public Function GetCode()
       If int(YSvoid.Format_Mid_Num(7))=0 Then
          GetCode="<font class=red2>验证码机制已关闭</font>"
          Exit function
        End If
		Dim test
		On Error Resume Next
		Set test=Server.CreateObject("Adodb.Stream")
		Set test=Nothing
		If Err Then
			Dim zNum
			Randomize timer
			zNum = cint(8999*Rnd+1000)
			Session("GetCode") = zNum
			GetCode="<input type=text class=txt name=CodeStr maxlength=4 size=4>&nbsp;"&Session("GetCode")		
		Else
			GetCode="<input type=text class=txt name=CodeStr maxlength=4 size=4>&nbsp;<img src='"&YSvoid.Web_Dir&"Include/YSvoid_GetCode.asp' align=absMiddle title= '验证码，看不清楚。请点击刷新验证码！' onclick=""this.src='"&YSvoid.Web_Dir&"Include/YSvoid_GetCode.asp'"">"		
		End If
	End Function

	'检查验证码是否正确
	Public Function CodeIsTrue()
       If int(YSvoid.Format_Mid_Num(7))=0 Then
          CodeIsTrue=True
          Exit function
        End If
		Dim CodeStr
		CodeStr=Lcase(Trim(Request("CodeStr")))
		If CStr(Session("GetCode"))=CStr(CodeStr) And CodeStr<>""  Then
			CodeIsTrue=True
			Session("GetCode")=empty
		Else
			CodeIsTrue=False
			Session("GetCode")=empty
		End If	
	End Function

	'//* 随机数 *//
	'-- 随机N位数字
	Public Function Rand_Num(ByVal rNum)
		Dim ri,rmax,rmin,rndnum
		rmax=10^(rNum)-1
		rmin=10^(rNum-1)
		Randomize
		rndnum=Int((rmax-rmin+1)*rnd)+rmin
		For ri=1 To rnum-Len(rndnum)
			rndnum="0"&rndnum
		Next
		Rand_Num=rndnum
	End Function

	'-- 随机文件名
	Public Function Rand_File(ByVal LeftVar)
		Dim temp1
		temp1=""
		If Is_Null(LeftVar)<>"" Then temp1=Left(LeftVar,1)
		Rand_File=temp1&Time_Type(now_time,11)&Rand_Num(2)
	End Function

	'//* 外部发言 *//
	'-- 判断发言是否来自外部
	Public Function Post_Chk()
		Dim server_v1,server_v2
		server_v1=Request.ServerVariables("http_referer")
		server_v2=Request.ServerVariables("server_name")
		If server_v1<>"" Then
			'mid(server_v1,8,len(server_v2))=server_v2
			server_v2="http://"&server_v2
			If Left(server_v1,Len(server_v2))=server_v2 Then
				Post_Chk=True
				Exit Function
			End If
		End If
		Post_Chk=False
	End Function

	Public Function Chk()
		Chk=False
		If Code_Form("chk")="yes" Then
			Chk=Post_Chk()
		End If
	End Function

	'//* 用户IP *//
	'-- IP SYS
	'-- 0=ip，1=sys
	Public Function Ip_Sys(ByVal St)
		If St=1 Then
			Ip_Sys=Request.ServerVariables("http_user_agent")
			Exit Function
		End If
		Dim userip,userip2
		userip=Request.ServerVariables("http_x_forwarded_for")
		userip2=Request.ServerVariables("remote_addr")
		If Instr(userip,",")>0 Then userip=Left(userip,Instr(userip,",")-1)
		If Instr(userip2,",")>0 Then userip2=Left(userip2,Instr(userip2,",")-1)
		If userip="" Then
			Ip_Sys=userip2
		Else
			Ip_Sys=userip
		End If
	End Function
	
	Public Sub Out_Msg(ByVal strMsg,ByVal strUrl)
		Response.Write "<Script Language=""JavaScript"">alert('"&strMsg&"');location.href='"&strUrl&"';</Script>"
		Response.end
	End Sub
	
	Public Function GetSize(size,unit)
		if isEmpty(size) or Not Isnumeric(size) then Exit Function
		size=CheckUnit(size,unit)
 		if size>1024 then
 			size=(size/1024)
 			getsize=formatnumber(size,2) & " MB"
		else
			getsize=size & " KB"
			Exit Function
 		end if
 		if size>1024 then
 			size=(size/1024)
 			getsize=formatnumber(size,2) & " GB"
 		end if
	End Function
	
	Public Function CheckUnit(size,unit)
		Select Case Lcase(Unit)
		Case "b"
			CheckUnit=formatnumber(size/1024,2)
		Case "k"
			CheckUnit=size
		Case "m"
			CheckUnit=(size*1024)
		Case "g"
			CheckUnit=(size*1024*1024)
		Case Else
			CheckUnit=size
		End Select
	End Function

	'//* JS输出 *//
	'-- 0 src  1 code
	Public Function Js_Put(ByVal j_var,ByVal j_type)
		Select Case Cint(j_type)
		Case 0
			Js_Put=VbCrlf&"<Script Language=""JavaScript"" src='"&j_var&"'></Script>"
		Case 1
			Js_Put=VbCrlf&"<Script Language=""JavaScript"">" & _
				   VbCrlf& j_var & _
				   VbCrlf&"</Script>"
		Case 2
			Js_Put=VbCrlf&"<Script Language=""JavaScript"">"
		Case 3
			Js_Put=VbCrlf&"</Script>"
		End Select
	End Function

	'//* 数字处理 *//
	Public Function Fix_Num(ByVal f_num,ByVal f_type)
		Dim fnum:fnum=f_num
		If Not(IsNumeric(fnum)) Then Fix_Num=0:Exit Function
		Select Case f_type
		Case 0
			Fix_Num=Int(fnum)
		Case 1
			If Instr(fnum,".")=0 Then
				Fix_Num=fnum
			Else
				Fix_Num=Int(fnum)+1
			End If
		Case 2
			If Instr(fnum,".")=0 Then
				Fix_Num=fnum
			Else
				If Cint(Mid(fnum,Instr(fnum,".")+1,1))>=5 Then
					Fix_Num=Int(fnum)+1
				Else
					Fix_Num=Int(fnum)
				End If
			End If
		End Select
	End Function

	'//* COOKIES、CACHE、SESSION操作部分开始 *//
	'-----------Cookies操作-------------
	Public Sub Cookies_Set(MyCookieName,CookieStr)
		Response.Cookies(Web_Cookies)(MyCookieName)=CookieStr
	End Sub

	Public Function Cookies_Get(MyCookieName)
		Cookies_Get=Trim(Request.Cookies(Web_Cookies)(MyCookieName))
	End Function

	Public Sub Cookies_Del(MyCookieName)
		Response.Cookies(Web_Cookies)(MyCookieName)=Empty
	End Sub

	Public Sub Cookies_Delele()
		Dim Cookies_Name
		For Each Cookies_Name In Request.Cookies(Web_Cookies)
			Response.Cookies(Web_Cookies)(Cookies_Name)=Empty
		Next
	End Sub
	'-----------Session操作-------------
	Public Sub Session_Set(MySessionName,SessionStr)
		Session(Web_Cookies&"_"&MySessionName )=SessionStr
	End Sub

	Public Function Session_Get(MySessionName)
		Session_get=Session(Web_Cookies&"_"&MySessionName )
	End Function

	Public Sub Session_Del(MySessionName)
		Session(Web_Cookies&"_"&MySessionName )=Empty
	End Sub

	Public Sub Session_Delele()
		Session.Abandon
	End Sub
	'------------缓存操作-------------

	Public Sub Cache_Del(MyCaheName)
		MyCaheName = LCase(MyCaheName)
		Application.Lock
		Application.Contents.Remove(Web_Cookies&"_"&MyCaheName)
		Application.UnLock
	End Sub

	Public Sub Cache_Delete()
		Dim Cacheobj
		For Each Cacheobj in Application.Contents
			If CStr(Left(Cacheobj,Len(Web_Cookies)+1))=CStr(Web_Cookies&"_") Then
				Application.Lock
				Application.Contents.Remove(Cacheobj)
				Application.UnLock
			End If
		Next
	End Sub

	Public Function Cache_Chk()
		Cache_Chk=True
		Cache_Data=Application(Web_Cookies & "_" & LocalCacheName)
		If Not IsArray(Cache_Data) Then Exit Function
		If Not IsDate(Cache_Data(1)) Then Exit Function
		If DateDiff("s",CDate(Cache_Data(1)),Now()) < (60*Reloadtime) Then Cache_Chk=False
	End Function

	Public Sub Cache_Set(MyCaheName,CaheStr)
		Name=MyCaheName
		Value=CaheStr
	End Sub

	Public Function Cache_Get(MyCaheName)
		Name=Lcase(MyCaheName)
		If Cache_Chk() Then
			Select Case Lcase(MyCaheName)
			Case "web_menu"
				MenuToCache(0)
			Case "admin_menu"
				MenuToCache(1)
			Case "user_group"
				UserGroup_To_Cache()
			Case "mainstyle_list"
				MainSkinList()
			Case "needtoupdate"
				Value=""
			Case "YSvoid_online","YSvoid_useronline"
				Online_Init()
			Case "forum_class"
		        Forum_Class_To_Cache()
			Case "forum_istops"
				Value=""
			Case "forum_istop_"&Forum_ID
				Value=""
			Case "forum_istopz_"&zid
				Value=""
			Case "forum_cast"
				Value=""
			Case "forum_class_select"
				Value=""
			End Select
		End If
		Name=Lcase(MyCaheName)
		Cache_Get=Value
	End Function

    '//* 发布信息间隔限制 *//
    Public Function Put_Space(ByVal Types,ByVal Num)
	  Dim pt_var,pt_num
	  Put_Space=True
	  pt_var=YSvoid.Session_Get(Types)
	  If IsDate(pt_var) Then
		pt_num=DateDiff("s",pt_var,YSvoid.now_time)
		If Int(Num)>0 And Int(pt_num)<=Int(Num) Then
			Put_Space=False
		End If
	  End If
	  YSvoid.Session_Set Types,YSvoid.now_time
    End Function

	Public Function YSvoid_encrypt(encrypt_var)
        if encrypt_var="" or isnull(encrypt_var) then YSvoid.YSvoid_encrypt="":exit function
        YSvoid_encrypt=encrypt_var
    End Function

    Public Function YSvoid_decrypt(decrypt_var)
        if decrypt_var="" or isnull(decrypt_var) then YSvoid.YSvoid_decrypt="":exit function
        YSvoid_decrypt=decrypt_var
     End Function

    Public Function ent_encrypt(vvar)
         ent_encrypt="YSvoid"&vvar
    End Function

    Public Function ent_unencrypt(vvar)
     ent_unencrypt=right(vvar,len(vvar)-2)
     End Function

	Public Function Img(ByVal vImg)
		Img="<img src='"&web_dir_skin&"small/"&vImg&".gif' border=0>"
	End Function

	Public Function Max_ID(ByVal vData,ByVal vID,ByVal vWhere)
		Dim oRs
		Max_ID=0
		If vWhere<>"" Then vWhere="Where "&vWhere
		Set oRs=Exec("Select Top 1 "&vID&" From "&vData&" "&vWhere&" Order By "&vID&" Desc",1)
		If Not (oRs.Eof And oRs.Bof) Then Max_ID=oRs(vID)
		oRs.Close:Set oRs=Nothing
	End Function

  Private Sub UserGroup_To_Cache()
		Dim oSQL, oRs
		oSQL = "Select GroupPower,GroupName,GroupType,GroupGrade,GroupConfig,FontColor,GlowColor,GroupEmoney From YSvoid_UserGroup Order By GroupOrder"
		Set oRs = Exec(oSQL,1)
		If Not (oRs.Eof And oRs.Bof) Then
			Value = oRs.GetRows(-1)
			oRs.Close:Set oRs = Nothing
			Exit Sub
		End If
		oRs.Close:Set oRs = Nothing
		Call ClsErr("用户组初始化失败！")
	End Sub

	Public Sub Forum_Class_To_Cache()
		Dim oSQL, oRs
		oSQL="Select Class_ID,Forum_ID,Forum_Name,Forum_F,Forum_N,Forum_W,Forum_Star,Forum_Parent,Forum_Power,Forum_Type,Forum_Pro,Forum_User,Forum_Config,Forum_New_Num,Forum_Special,Forum_Rule From YSvoid_BbsForum Where Forum_Hidden=0 Order By Class_ID,Forum_Order,Forum_Star"
		Set oRs = Exec(oSQL,1)
		If Not (ORs.Eof And ORs.Bof) Then
			Value=ORs.GetRows(-1)
			ORs.Close:Set oRs = Nothing
			Exit Sub
		End If
		'Value=""
		oRs.Close:Set oRs = Nothing
		Call ClsErr("论坛分类初始化失败！")
	End Sub

	'//* 网站人数统计类 *//
	Public YSvoid_Online							'网站在线人数
	Public YSvoid_UserOnline						'用户登陆人数
	Public YSvoid_GuestOnline						'游客人数
	'--- 统计人数初始化
	Private Sub Online_Init()
		Name = "YSvoid_Online"
		Reloadtime = 60
		If Cache_Chk() Then ReLoadOnlineNum
		Name = "YSvoid_Online"
		YSvoid_Online = Value
		Name = "YSvoid_UserOnline"
		If Cache_Chk() Then ReLoadOnlineNum
		YSvoid_UserOnline = Value
		If YSvoid_Online < 0  Or YSvoid_UserOnline < 0 Or YSvoid_UserOnline > YSvoid_Online Then ReLoadOnlineNum
		YSvoid_GuestOnline = YSvoid_Online - YSvoid_UserOnline
		Reloadtime = 14400
	End Sub

	Public Sub YSvoid_DelOnlineNum(ByVal vName)
		If vName="" Then Exit Sub
		Call Exec("delete from YSvoid_UserLogin where l_username='"&vName&"'",0)
		YSvoid_Online=YSvoid_Online-1
		Cache_Set "YSvoid_Online",YSvoid_Online
	End Sub

	'--- 更新在线人数，删除过期用户
	Public Sub UpdateOnline()
		'Call Online_Init()
		Dim SQL,SQL1
		Dim TempNum,TempNum1
		Reloadtime = 60
		Name="Online_DelTime"
		If Cache_Chk() Then Value=now_time
		If DateDiff("s",Value,now_time) > Clng(YSvoid.Format_Mid_Num(14))*60 Or DateDiff("s",Value,now_time) = 0 Then
			Value=now_time
				sql="delete from YSvoid_UserLogin where l_type=0 and datediff('s',l_tim_end,'"&now_time&"')>"&YSvoid.Format_Mid_Num(14)&"*60"
				sql1="delete from YSvoid_UserLogin where l_type=1 and datediff('s',l_tim_end,'"&now_time&"')>"&YSvoid.Format_Mid_Num(14)&"*60"
			call exec("",-1)
			Conn.Execute SQL,TempNum
			Conn.Execute SQL1,TempNum1
			'如果删除客人数大于0，则应该更新总数
			If TempNum>0 Then
				'更新缓存总在线数据
				YSvoid_Online = YSvoid_Online - TempNum
				YSvoid_GuestOnline = YSvoid_GuestOnline - TempNum
			End If
			'如果删除用户数大于0，则应该更新总数和用户数
			If TempNum1>0 Or TempNum>0 Then
				'更新缓存总在线数据
				YSvoid_Online = YSvoid_Online - TempNum1
				YSvoid_UserOnline = YSvoid_UserOnline - TempNum1
			End If
			Name="YSvoid_Online"
			Value=YSvoid_Online
			'更新缓存总用户在线数据
			Name="YSvoid_UserOnline"
			Value=YSvoid_UserOnline
			'YSvoid_Online = YSvoid_Online - TempNum1
		End If
		Reloadtime = 14400
	End Sub

	'--- 刷新在线数据缓存
	Public Sub ReLoadOnlineNum
		Dim oRs
		Name="YSvoid_Online"
		Set oRs=Exec("Select Count(l_id) From YSvoid_UserLogin",1)
		Value = oRs(0)
		oRs.Close
		YSvoid_Online = Value
		Name="YSvoid_UserOnline"
		Set oRs=Exec("Select Count(l_id) From YSvoid_UserLogin Where l_type=1",1)
		Value=oRs(0)
		YSvoid_UserOnline = Value
		oRs.Close:Set oRs=Nothing
	End Sub

	'--- 查询在某版面的在线总数
	Public Property Get Forum_Online
		Forum_Online=Forum_UserOnline+Forum_GuestOnline
	End Property

	Public Property Get Forum_GuestOnline
		Dim l_GuestOnline, oRs
		l_GuestOnline=-1
		If Forum_ID>0 Then
			Set oRs=Exec("Select Count(l_id) From YSvoid_UserLogin Where l_forum_id="&Forum_ID&" And l_type=0",1)
			l_GuestOnline=oRs(0)
			oRs.Close:Set oRs=Nothing
		End If
		Forum_GuestOnline=l_GuestOnline
	End Property

	Public Property Get Forum_UserOnline
		Dim l_Online, oRs
		l_Online=-1
		If Forum_ID>0 Then
			Set oRs=Exec("Select Count(l_id) From YSvoid_UserLogin Where l_forum_id="&Forum_ID&" And l_type=1",1)
			l_Online=oRs(0)
			oRs.Close:Set oRs=Nothing
		End If
		Forum_UserOnline=l_Online
	End Property

'--- 频道数据统计
   Public Function Channel_Count(iChannelID)
	    Dim iCount, iHits
		If iChannelID=5 Then
	    SQL="Select Count(*) From YSvoid_BbsTopic"
		Else
		SQL="Select Count(*) From ["&YSvoid.ChannelTable&"] Where Hidden=1 And ChannelID="&iChannelID
		End If
		Set Rs=YSvoid.Exec(SQL,1)
		iCount=Rs(0)
		Rs.Close
		If Not YSvoid.Is_Int(iCount) Then iCount=0
		If iChannelID=5 Then
		  SQL="Select Count(*) From YSvoid_BbsData"
		Else
		  SQL="Select Sum(Counter) From ["&YSvoid.ChannelTable&"] Where Hidden=1 And ChannelID="&iChannelID
		End If
		Set Rs=YSvoid.Exec(SQL,1)
		iHits=Rs(0)
		Rs.Close
		If Not YSvoid.Is_Int(iHits) Then iHits=0
		Call YSvoid.Exec("Update [YSvoid_Channel] Set ChannelCount="&iCount&",ChannelHits="&iHits&" Where ChannelID="&iChannelID,0)
		Set Rs=Nothing
		Call Cache_Del("web_menu")
   End Function


'--- 频道服务器读取
   Public Function Server_Url(oChannelid,ser_id,ser_url)
		dim s_url
  			sql="select S_Url from YSvoid_Server where Channelid="&Channelid&" and S_ID="&ser_id
  			set rs=YSvoid.exec(sql,1)
  			if not rs.eof then
				S_Url=rs(0)
				if right(S_Url,1)<>"/" then s_url=s_url&"/"
				Server_Url=s_url&ser_url
			else
				Server_Url=ser_url
			end if
			rs.close
  	end Function


  '---UBB---
	Public Function UbbCode(str)
	  Dim YSvoidUbb
		if Is_Null(str)<>"" then
			Set YSvoidUbb=New Cls_UbbCode
			UbbCode=YSvoidUbb.UbbCode(str)
			Set YSvoidUbb=Nothing
		end if
	End Function
End Class

%>