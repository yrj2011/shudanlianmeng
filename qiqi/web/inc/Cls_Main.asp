<%

Class Cls_YSvoidMain
  public web_vt,web_name,web_url,Web_RealURL,web_email,web_now_dir,web_urls,web_dir,web_cookies,web_skin,web_dir_skin,web_unit,now_time,today_time,timer_start,pro_edition,num_rs

	Public YSvoid_Info			'ϵͳһЩͳ����Ϣ
	Public Web_Common
	Public IsHealth
	Public SuperAdmins, AdminFolder, Web_Upload, Web_UploadFilters, Web_BadWords, Web_LockIP
	Public Web_LockTime, Web_Boy, Web_Girl, Web_Index
	Public Web_AdminCodes ,Web_Number ,Web_UploadType ,Web_MailServer ,Web_RegBadWords, Web_RegUserName
	Public Web_CopyRight, Web_Closer, Web_Mode
	Public Web_UserLevel, Web_PerPay, Web_Medal, Web_Gbook, Web_ForumData, Web_Payment
	Public IsDeBug
	Public Reloadtime,CacheName
	Public Fso_Sys_Var										'FSO ������
	Private LocalCacheName,Cache_Data

	Public IsConn											'Conn �Ƿ���

	Private WebConfig										'��վ�������ã�����
	'ģ��˽�б���
	Private iMainSkinID
	Private iMainName							'��ǰģ������
	Private iMainIndex							'��ҳģ��
	Private iMainWhole							'ģ����
	Private iMainOther, iMainOtherName			'����ģ�塢��ǩ����
	Private iMainStrer, iMainStrerName			'ȫ���ַ�������ǩ����
	Private iPageStrer, iPageStrerName			'Ƶ���ַ�������ǩ����
	Private iPageOther, iPageOtherName			'Ƶ������ģ�塢��ǩ����
	Private iMainWidth							'���ֵ
	Private iMainPicPath						'ͼƬ·��
	Private iMainCssPath						'Css·��
	Private iMainConfig							'����ж�
	Private PutHtml

	Public total_counter									'ͳ������
	Public IsView'�ǵ�Ҫɾ����
	Public PageModTrue										'�ж���Ŀ�Ƿ�����ģ��

	Public ChannelTrue
	Public ChannelName,ChannelTit,ChannelDir,ChannelTable,ChannelUnit,ChannelListNum,ChannelPosition
	Public ChannelSetup,ClassDepth,ChannelPutType,ChannelCuteNum,ChannelLeftCuteNum,ChannelCount,ChannelHits,ChannelOption,ChannelReward,ChannelDates,ChannelRss,ChannelPower,ChannelPHidden
	Public ChannelUpType,ChannelUpMax,ChannelUpDayNum,ChannelUpFileType,ChannelUploadDir
	Public ChannelPicWidth, ChannelPicHeight
	Public ReviewTrue, SpecialTrue, KeyWordTrue, ErrTrue, CastTrue
	Public ChannelSpecial, ChannelKeyWord, ChannelCast

	'//* ���ʼ�� *//
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

    '//* ����� *//
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


	'//* ����ϵͳ�������� *//
	Private Sub Cms_InfoInit()
		Dim Tmpstr
		'0-ע��������1-���ע���û���2-����������3-�������������4-��߷���ʱ�䡢5-��վ��ʼʱ��
		'6-����ʱ�䡢7-����������8-���ò�����9-�����û���10-��վFso��11-��̨Ŀ¼��12-�ϴ�Ŀ¼
		'13-�������ַ���14-����IP��15-�û��ȼ���16-Ԥ���ѡ�17-���Բ�����18-��ǰ��̳���ݱ���Ϣ
		'19-֧��������20-�ϴ�����
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
		'--- ÿ��ˢ��һ��ͳ����Ϣ
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


	'--- ���������ñ��ֻ������飬��ڣ��������ݡ�����λ��
	Public Sub ReloadInfo(MyValue,N)
		YSvoid_Info(N,0) = MyValue
		Name = "YSvoid_info"
		value = YSvoid_Info
		YSvoid_Info = value
	End Sub

	'--- ÿ�ռ򵥳�ʼ��
	Private Sub Day_Initialize()
		Dim dTim
		dTim = Time_Type(Now_Time,4)
		Call Exec("Update YSvoid_Configs Set Num_New=0,Today_Tim='"&dTim&"'",0)
		Call Exec("Update YSvoid_BbsForum Set Forum_New_Num=0",0)
		Call Exec("Update YSvoid_UserData Set abate=1 Where estate=1 And otim<"&dTim,0)
		ReloadInfo 0,7
		ReloadInfo dTim,6
	End Sub

	'--- ȡ��������
	Public Property Get Cache_Num
		Cache_Num=Application.Contents.count
	End Property

	'--- ȡ�û�������
	Public Property Get UserGroupLen
		UserGroupLen=0
		Name="user_group"
		If Cache_Chk() Then UserGroup_To_Cache()
		If IsArray(Value) Then
			UserGroupLen=Ubound(Value,2)
		End If
	End Property

	'--- �����ǰһҳURL������Encode
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

	'�ж��û��Ƿ�����
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

	'�����û����ϻ���(�����û���,�Ƿ���Ҫ���)[0=�����,ֻ������,1=��Ҫ���]
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

	'--- ��̳ID ---
	Public Property Get Forum_ID
		Forum_ID=Code_id("fid")
	End Property

	'--- ҳ��ִ��ʱ�� ---
	Public Property Get Exec_Tim
		Exec_Tim=FormatNumber((Timer()-Timer_Start),3)
		If Left(Exec_Tim,1)="." Then Exec_Tim="0"&Exec_Tim
	End Property

	'//* ģ��������ֿ�ʼ *//
	'-----------ģ���б�---------------
	Private Sub MainSkinList()
		Dim oSQL, oRs
		oSQL = "Select TemplateID,TemplateName From [YSvoid_Template] Where ChannelID=0"
		Set oRs = Exec(oSQL,1)
		Value = oRs.GetRows(-1)
		oRs.Close:Set oRs=Nothing
	End Sub

	'-----------ȡĬ��ģ�� ID-----------
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
			ClsErr("ģ�������ǿյģ�������ӡ�")
		End If
	End Sub

	'-----------ȡģ������---------
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
			'�������
			If tChannelID = 0 Then
				Response.Redirect "Skins.asp?Action="&iMainSkinID
			Else
				ClsErr("Ƶ��ģ�������ǿյģ�������ӡ�")
			End If
		End If
	End Sub

	'--------����ģ������------------
	Private Sub MainMod_Load()
		If IsDeBug = 1 Then On Error Resume Next'�ݴ���롣
		Dim AllStr, Tmpstr, iSkinID, iCookiesSid
		'-----���Ĭ��ģ��ID
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
		'-----ȡģ������
		Name="MainStyle"&iSkinID
		If Cache_Chk() Then Call TemplateCache(0,iSkinID)
		AllStr = Value
		If Not IsArray(AllStr) Then
			ClsErr("Ƶ��ģ�����ݵ���������ϵ����Ա��")
			Exit Sub
		End If
		If Ubound(AllStr,1)<>5 Then
			ClsErr("Ƶ��ģ�����ݵ���������ϵ����Ա��")
			Exit Sub
		End If
		'-----ȡCSS·��
		iMainCssPath = YSvoid.Web_Dir&"Skin/Skin_"&iSkinID&".css"
		iMainCssPath = "<link type=""text/css"" href="""&iMainCssPath&""" rel=stylesheet>"
		'-----ȡģ������
		iMainName = AllStr(0,0)
		'-----ȡ��ҳģ��
		iMainIndex = AllStr(1,0)
		'-----ȡģ����
		iMainWhole = Split(AllStr(2,0),"|||")
		'-----ȡ����ģ��
		Tmpstr = Split(AllStr(3,0),"@@@")(0)
		iMainOther = Split(Tmpstr,"|||")
		Tmpstr = Split(AllStr(3,0),"@@@")(1)
		iMainOtherName = Split(Tmpstr,"|||")
		'-----ȡ�ַ���
		Tmpstr = Split(AllStr(4,0),"@@@")(0)
		iMainStrer = Split(Tmpstr,"|||")
		Tmpstr = Split(AllStr(4,0),"@@@")(1)
		iMainStrerName = Split(Tmpstr,"|||")
		'-----ȡģ��ҳ����
		iMainWidth = iMainWhole(12)
		'-----ȡͼƬ·��
		iMainPicPath = Split(AllStr(5,0),"|||")(0)
		Erase AllStr
		web_dir_skin = YSvoid.Web_Dir&"Skin/"&iMainPicPath&"/"
	End Sub

	'--- �ж���Ŀģ���Ƿ����벢��ֵ
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
		'-----ȡģ���ܣ�����ʹ��ȫ�ֿ��
		If AllStr(0,0) = True Then
			iMainWhole = Split(AllStr(1,0),"|||")
			'-----ȡģ��ҳ����
			iMainWidth = iMainWhole(12)
		End If
		'-----ȡ����ģ��
		Tmpstr = Split(AllStr(2,0),"@@@")(0)
		iPageOther = Split(Tmpstr,"|||")
		Tmpstr = Split(AllStr(2,0),"@@@")(1)
		iPageOtherName = Split(Tmpstr,"|||")
		'-----ȡ�ַ����������ڶ���������
		Tmpstr = Split(AllStr(3,0),"@@@")(0)
		iPageStrer = Split(Tmpstr,"|||")
		Tmpstr = Split(AllStr(3,0),"@@@")(1)
		iPageStrerName = Split(Tmpstr,"|||")
		iPageCss = AllStr(4,0)
		If iPageCss <> "" Then
			iMainCssPath = iMainCssPath & Vbcrlf & "<style>"
			iMainCssPath = iMainCssPath & Vbcrlf & "<!-- Ƶ�� CSS "
			iMainCssPath = iMainCssPath & Vbcrlf & iPageCss
			iMainCssPath = iMainCssPath & Vbcrlf & "-->"
			iMainCssPath = iMainCssPath & Vbcrlf & "</style>"
		End If
		Erase AllStr
		Channel_TemplateLoad = True
	End Function

	'--- ģ���г��������滻 ---
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

	'--- ȡȫ��ģ��(���鷽ʽ) ---
	Public Sub HtmlWhole(Mnum)
		PutHtml=iMainWhole(Mnum)
	End Sub

	'--- ȡȫ��ģ��(������ʽ) ---
	Public Sub HtmlMain(Mname)
		PutHtml=HtmlStream(Mname,iMainOtherName,iMainOther)
	End Sub

	'--- ȡ��ҳģ��(���鷽ʽ) ---
	Public Sub HtmlPageNum(Mnum)
		PutHtml=iPageOther(Mnum)
	End Sub

	'--- ȡ��ҳģ��(������ʽ) ---
	Public Sub HtmlPage(Mname)
		PutHtml=HtmlStream(Mname,iPageOtherName,iPageOther)
	End Sub

	'--- ������ʽ��ȡHTML������
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

	'--- ģ���ǩ�滻 ---
	Public Sub HtmlRcod(Mod_Str,View_Str)
		If IsDeBug=0 Then On Error Resume Next
		'Mod_Str=LCase(Mod_Str)
		PutHtml=Replace(PutHtml,"{$"&Mod_Str&"}",View_Str)
	End Sub

	'--- ģ���ӡ ---
	Public Sub HtmlView(Knum)
		Response.Write ReHtml(PutHtml,0) & VbCrlf
		'Response.Write ReHtml(PutHtml,0)
		If Knum=1 Then Response.Write ModGet("ukong")
	End Sub

	'--- ģ�帳ֵ ---
	Public Function HtmlGet(Knum)
		HtmlGet=ReHtml(PutHtml,0) & VbCrlf
		'HtmlGet=ReHtml(PutHtml,0)
		If Knum=1 Then HtmlGet=HtmlGet & ModGet("ukong")
	End Function
	'//* ģ��������� *//

	'//* ��վ���˵�Ƶ����Ϣ����ʼ *//
	'//* mType:( 0=ǰ̨��1=��̨ ) *//
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

	'--- ȡƵ��������Ϣ
	Public Property Get Channel_Menu
		Channel_Menu=Cache_Get("web_menu")
	End Property
	'//* ��վ���˵�Ƶ����Ϣ������� *//

	'//* Ƶ����Ϣ���� *//
	'--- Ƶ����ʼ������
	Public Sub Channel_Initialize(ByVal nChannelID)
		Dim ChannelDim, Cii, iUpFileType
		If nChannelID =< 0 Then Exit Sub
		'--- ����ǰ̨Ƶ������
		ChannelDim = Channel_Menu
		'--- ���в˵��������ݲ����и�ֵ
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

	'--- ��ȡĳ��Ƶ����ĳ������ֵ
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

	'--- ��ȡƵ������ֵ
	'--- 0-����Ȩ��1-���ࡢ2-���ۡ�3-ר�⡢4-�ؼ��֡�5-����
	'--- 6-RSS��7-�Ű���ʽ��8-�����ַ���9-ǰ̨�༭����10-��̨�༭����13-����
	Public Function Get_ChannelSetup(ByVal cSetup,ByVal sNum)
		Get_ChannelSetup = 0
		sNum = sNum + 1
		If Len(cSetup) <> 20 Or sNum > 20 Then Exit Function
		Get_ChannelSetup = Int(Mid(cSetup,sNum,1))
	End Function

	'//* Ƶ��ר�⡢�ؼ��֡����洦�� *//
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

	'//* ģ�建��������ֿ�ʼ *//
	'--- ת����������Сд
	Public Property Let Name(ByVal vNewValue)
		LocalCacheName=LCase(vNewValue)
	End Property

	'--- ���渳ֵ�����鷽ʽ
	Public Property Let Value(ByVal vNewValue)
		If LocalCacheName<>"" Then
			ReDim Cache_Data(2)
			Cache_Data(0)=vNewValue
			Cache_Data(1)=Now()
			Application.Lock
			Application(Web_Cookies & "_" & LocalCacheName) = Cache_Data
			Application.unLock
		Else
			ClsErr("��������δ���� �� �����е��ظ������������ã�")
		End If
	End Property

	'--- ����ȡֵ������0
	Public Property Get Value()
		If LocalCacheName<>"" Then
			Cache_Data=Application(Web_Cookies & "_" & LocalCacheName)
			If IsArray(Cache_Data) Then
				Value=Cache_Data(0)
			Else
				ClsErr("���� ("&LocalCacheName&") û������ ��")
			End If
		Else
			ClsErr("��������δ���� �� �����е��ظ������������ã�")
		End If
	End Property

	'//* ��ӡ�����ҳ�� *//
	Private Sub ClsErr(ClsStr)
		Response.Clear
		Response.Write ClsStr
		Response.End
	End Sub


	'//* ���ݿ�������ֿ�ʼ *//
	Public Function Exec(Esql,Etype)
		If IsConn=False Then Call Conn_Open()
		If IsDeBug=1 Then On Error Resume Next
		Select Case Etype
		Case 0
			Conn.Execute(Esql)
		Case 1
			Set Exec=Conn.Execute(Esql)
		End Select
		If IsView=1 Then Response.Write "������䣺"&esql&"<br>"
		'If IsView=1 Then Call ClsErr("������䣺"&esql&"<br>")
		If Err And IsDeBug=1 Then
			Err.Clear
			Response.Write "������䣺"&esql &"<br>"
			Response.End
		End If
		Num_Rs=Num_Rs+1
	End Function

	Public Function Exe_Conn(Ers,Esql,Etype)
		If IsConn=False Then Call Conn_Open()
		If IsDeBug=1 Then On Error Resume Next
		Ers.Open Esql,Conn,1,Etype
		If IsView=1 Then Response.Write "������䣺"&esql&"<br>"
		If Err And IsDeBug=1 Then
			Err.Clear
			Response.Write "������䣺"&esql&"<br>"
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

    '����:SqlServer(97-2000) to Access(97-2000)
    '����:Sql,���ݿ�����(ACCESS,SQLSERVER)
    '˵��:
    Private Function Sql_To_Access(ByVal vNewSql)
        Dim regEx, Matches, Match, vSql
        vSql=vNewSql
        '�����������
        Set regEx = New RegExp
        regEx.IgnoreCase = True
        regEx.Global = True
        regEx.MultiLine = True
        'ת:GetDate()
        regEx.Pattern = "(?=[^']?)GETDATE\(\)(?=[^']?)"
        vSql = regEx.Replace(vSql,"NOW()")
        'ת:UPPER()
        regEx.Pattern = "(?=[^']?)UPPER\([\s]?(.+?)[\s]?\)(?=[^']?)"
        vSql = regEx.Replace(vSql,"UCASE($1)")
        'ת:���ڱ�ʾ��ʽ
        '˵��:ʱ���ʽ������2004-23-23 11:11:10 ��׼��ʽ
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
        'ת:Insert����
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

  '-------------------------------------��ַ����--------------------------------------
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

  '-----------------------------------ʱ���ʽ����------------------------------------
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
    case 2	'��(4)-��-�� ʱ:��:��
      time_type=d_year&"��"&d_month&"��"&d_day&"�� "&d_hour&"ʱ"&d_minute&"��"&d_second&"��"
    case 3	'10-10 23:45
      time_type=d_month&"-"&d_day&" "&d_hour&":"&d_minute
    case 4	'2003-10-10
      time_type=d_year&"-"&d_month&"-"&d_day
    case 5	'2003��10��10��
      time_type=d_year&"��"&d_month&"��"&d_day&"��"
    case 6	'10-10
      time_type=d_month&"-"&d_day
    case 7	'03-10-10
      time_type=right(d_year,len(d_year)-2)&"-"&d_month&"-"&d_day
    case else
      time_type=ttt
    end select
  End Function

	'//* �ַ������� *//
'-- �ж��ַ����Ƿ�Ϊ��
	Public Function Is_Null(ByVal sVal)
		Is_Null=Trim(sVal)
		If Is_Null="" Or IsNull(Is_Null) Then Is_Null=""
	End Function

	'-- �ж��Ƿ�Ϊ�������ͣ����ز���ֵ
	Public Function Is_Int(ByVal sVal)
		Is_Int=True
		If Is_Null(sVal)="" Or Not(IsNumeric(sVal)) Or Instr(sVal,".")>0 Then Is_Int=False
	End Function

	'-- �ж��Ƿ�Ϊ�������ͣ����ز���ֵ
	Public Function Is_Num(ByVal sVal)
		Is_Num=True
		If Is_Null(sVal)="" Or Not(IsNumeric(sVal)) Then Is_Num=False
	End Function

	'-- E-Mail�ж�
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

	'-- �ַ�����
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

	'-- ���� Form ֵ
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

	'-- ���� Query ֵ
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
 		Strer=Replace(Strer,"","&quot;")		'˫����
        Strer=Replace(strer,Chr(33),"&nbsp;")		'�ո�
		Code_Remark=Strer
	End Function

	'-- ���� ID ����ֵ
	Public Function Code_Id(ByVal sVal)
		Dim Strer
		Strer=Trim(Request.Querystring(sVal))
		If Not(Is_Int(Strer)) Then Strer=0
		Code_Id=Clng(Strer)
	End Function

	'-- ��������ֵ
	Public Function Code_Int(ByVal Strers,ByVal Sint)
		Dim Strer
		Strer=Code_Form(Strers)
		If Not(Is_Int(Strer)) Then
			Code_Int=Sint
			Exit Function
		End If
		Code_Int=Int(Strer)
	End Function

	'ȡ���ַ���ֵ
	'Strers:�ַ���
	'At    :����
	'Acut  :�ض��ַ�����
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
		Strer=Replace(Strer,Chr(39),"&#39;")		'������
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
		Strer=Code_Health(Strer)		'���˲�������Ϣ
		Strer=Replace(Strer,"\","\\")
		Strer=Replace(Strer,Chr(39),"&#39;")		'������
		'Strer=Replace(Strer,Chr(39),"\'")
		Strer=Replace(Strer,Chr(34),"&quot;")		'˫����
		'Strer=Replace(Strer,Chr(34),"\""")
		Select Case Tt
		Case 0
			'Strer=Replace(Strer,VbCrlf,"")			'�س�
			Strer=Replace(Strer,Chr(10),"")
			Strer=Replace(Strer,Chr(13),"")
		Case 1
			'Strer=Replace(Strer,VbCrlf,"\n")			'�س�
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
		Strer=Code_Health(Strer)		'���˲�������Ϣ
		If CuteNum>0 Then Strer=Cuted(Strer,CuteNum)
		Strer=Replace(Strer,"<","&lt;")
		Strer=Replace(Strer,">","&gt;")
		Strer=Replace(Strer,Chr(39),"&#39;")		'������
		Strer=Replace(strer,Chr(34),"&quot;")		'˫����
		Strer=Replace(strer,Chr(33),"&nbsp;")		'�ո�
		Select Case Chtype
		Case 1
			Strer=Replace(Strer,Chr(9),"&nbsp;")		'table
			'Strer=Replace(Strer,VbCrlf,"")			'�س�
            Strer=Replace(Strer,"{$download_dir}",Web_Dir&Web_Upload)
			Strer=Replace(Strer,Chr(10),"")
			Strer=Replace(Strer,Chr(13),"")
		Case 2
			Strer=Replace(Strer,Chr(9),"&nbsp;��&nbsp;")	'table
			'Strer=Replace(Strer,VbCrlf,"<br \>")		'�س�
			Strer=Replace(Strer,Chr(10),"<br \>")
			Strer=Replace(Strer,Chr(13),"")
		End Select
		Code_Html=Strer
	End Function

	'���˲��������ַ���
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

	    '���˷Ƿ���SQL�ַ�
	Public Function ReplaceBadChar(Byval strChar)
		strChar=replace(replace(strChar," ",""),"'","")
		strChar=replace(replace(strChar,".",""),"<","")
		strChar=replace(replace(strChar,")",""),"(","")
		strChar=replace(replace(strChar,"?",""),"*","")
		strChar=replace(replace(strChar,"/",""),"\","")
		ReplaceBadChar=replace(strChar,Chr(0),"")
	End Function
	
    '��ʶ�������ֵ�ͨ���ַ��ָ���(ȡָ�����ȵ��ַ���������ʾ�ã�OK
    'Types      Ϊ����Ĵ��ָ��ַ�
    'Num        Ϊ�ָ�ĳ�(�ֽ���������Ϊ���ֽ�),����ڼ���
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

    '��ʶ�������ֵ�left(),����Ϊ���ֽ� OK
    'Strers   Ϊԭ�ַ���
    'Num     Ϊ��Ҫ(��)��ȡ��λ��
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

    '��ʶ�������ֵ�len(),����Ϊ�� OK
    'Strers    Ϊ����ⳤ�ȵ��ַ���
    Public Function Str_Length(ByVal Strers)
        On Error Resume Next
        Dim Winnt_Chinese
        Winnt_Chinese = (Len("�й�") = 2)
        If Winnt_Chinese = True Then
            Dim LNum, TNum, CNum, i
            LNum = Len(Strers)
            TNum = LNum
            For i = 1 To LNum
                CNum = Asc(Mid(Strers, i, 1)) '�����ַ��Ĵ���ֵ
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

	'��ʾ��֤��
	Public Function GetCode()
       If int(YSvoid.Format_Mid_Num(7))=0 Then
          GetCode="<font class=red2>��֤������ѹر�</font>"
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
			GetCode="<input type=text class=txt name=CodeStr maxlength=4 size=4>&nbsp;<img src='"&YSvoid.Web_Dir&"Include/YSvoid_GetCode.asp' align=absMiddle title= '��֤�룬�������������ˢ����֤�룡' onclick=""this.src='"&YSvoid.Web_Dir&"Include/YSvoid_GetCode.asp'"">"		
		End If
	End Function

	'�����֤���Ƿ���ȷ
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

	'//* ����� *//
	'-- ���Nλ����
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

	'-- ����ļ���
	Public Function Rand_File(ByVal LeftVar)
		Dim temp1
		temp1=""
		If Is_Null(LeftVar)<>"" Then temp1=Left(LeftVar,1)
		Rand_File=temp1&Time_Type(now_time,11)&Rand_Num(2)
	End Function

	'//* �ⲿ���� *//
	'-- �жϷ����Ƿ������ⲿ
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

	'//* �û�IP *//
	'-- IP SYS
	'-- 0=ip��1=sys
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

	'//* JS��� *//
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

	'//* ���ִ��� *//
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

	'//* COOKIES��CACHE��SESSION�������ֿ�ʼ *//
	'-----------Cookies����-------------
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
	'-----------Session����-------------
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
	'------------�������-------------

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

    '//* ������Ϣ������� *//
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
		Call ClsErr("�û����ʼ��ʧ�ܣ�")
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
		Call ClsErr("��̳�����ʼ��ʧ�ܣ�")
	End Sub

	'//* ��վ����ͳ���� *//
	Public YSvoid_Online							'��վ��������
	Public YSvoid_UserOnline						'�û���½����
	Public YSvoid_GuestOnline						'�ο�����
	'--- ͳ��������ʼ��
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

	'--- ��������������ɾ�������û�
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
			'���ɾ������������0����Ӧ�ø�������
			If TempNum>0 Then
				'���»�������������
				YSvoid_Online = YSvoid_Online - TempNum
				YSvoid_GuestOnline = YSvoid_GuestOnline - TempNum
			End If
			'���ɾ���û�������0����Ӧ�ø����������û���
			If TempNum1>0 Or TempNum>0 Then
				'���»�������������
				YSvoid_Online = YSvoid_Online - TempNum1
				YSvoid_UserOnline = YSvoid_UserOnline - TempNum1
			End If
			Name="YSvoid_Online"
			Value=YSvoid_Online
			'���»������û���������
			Name="YSvoid_UserOnline"
			Value=YSvoid_UserOnline
			'YSvoid_Online = YSvoid_Online - TempNum1
		End If
		Reloadtime = 14400
	End Sub

	'--- ˢ���������ݻ���
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

	'--- ��ѯ��ĳ�������������
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

'--- Ƶ������ͳ��
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


'--- Ƶ����������ȡ
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