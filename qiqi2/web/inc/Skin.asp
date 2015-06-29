<!-- #Include File="Config.asp" -->
<!-- #Include File="Functions.asp" -->
<%

Dim m_channel,web_img_m,web_img_h
Dim login_is_abate,login_is_estate,login_is_otim

'// 用户IP信息
Dim user_ip
	user_ip = YSvoid.Ip_Sys(0)

'// 用户系统信息
Dim user_sys
	user_sys = YSvoid.Code_Var(Left(YSvoid.Ip_Sys(1),250))

'// Cookies支持
Dim Cookies_True
	Cookies_True = True
	If YSvoid.Cookies_Get("cookies_true")<>"yes" Then
		Response.Cookies(YSvoid.web_cookies).Path = YSvoid.Web_Dir
		YSvoid.Cookies_Set "cookies_true","yes"
		Cookies_True = False
	End If

'// Frame 显示/隐藏
Dim iFrame
	iFrame = False

'// 常用变量设置
Dim pic_w,pic_h,max_w,max_h,img_emoney,m_hei
  pic_w = YSvoid.Format_Mid_Num(22)
  If int(YSvoid.Format_Mid_Num(22))<>int(YSvoid.ChannelPicWidth) and int(YSvoid.ChannelPicWidth)>0 Then pic_w = YSvoid.ChannelPicWidth
  pic_h = YSvoid.Format_Mid_Num(23)
  If int(YSvoid.Format_Mid_Num(23))<>int(YSvoid.ChannelPicHeight) and int(YSvoid.ChannelPicHeight)>0 Then pic_h = YSvoid.ChannelPicHeight
  max_w = YSvoid.Format_Mid_Num(24)
  m_hei = YSvoid.Format_Mid_Num(13)
  img_emoney = "<img border=0 src='"&YSvoid.web_dir_skin&"small/emoney.gif' alt='"&YSvoid.web_unit&"' align=absmiddle>"

'// 用户是否登陆
Dim Is_Login_true
	Is_Login_true = False

'// 刷新限制
Dim Is_Refur_True
	Is_Refur_True = True

'// 网站计数类型
Dim Stamp_Num

'// 在线人数、在线用户数
Dim Online_Num,UserOnline_Num
	Online_Num = 0
	UserOnline_Num = 0

'// 网站是否开启
Call YSvoid_IsOpen()

'// 网站设置时间是否开启
Call YSvoid_TimOpen()

'// 用户 IP 是否被锁定
Call web_ip_shield()

'//* 网站标识 *//
Sub Web_Head_Tit()
	tit=Replace(tit,"---"," - ")
	tit=YSvoid.Code_Var(tit)
	tit_fir=YSvoid.Code_Var(tit_fir)
	YSvoid.HtmlWhole 0
	YSvoid.HtmlRcod "web_generator",YSvoid.Web_Dim(3)
	YSvoid.HtmlRcod "web_keywords",Replace(YSvoid.Web_Dim(4),"|",",")
	YSvoid.HtmlRcod "web_description",tit
	YSvoid.HtmlView(0)
	If iFrame Then Response.Write VbCrlf&YSvoid.ModGet("iframe")
End Sub

'================================================
'用 途：站内广告模块
'参 数：
'	aSort：广告名称
'	aType：是否留空
'================================================
Sub YSvoid_Ads(ByVal aSort,ByVal aType)
	If YSvoid.Is_Null(aSort)="" Then Exit Sub
	Response.Write YSvoid.Js_Put(YSvoid.Web_Dir&"tbmama/mama_"&aSort&".js",0)
	If aType = 1 Then
		Response.Write uKong
	End If
End Sub

Sub Channel_Loader()
	Dim SkinTrue
	Call YSvoid.Channel_Initialize(ChannelID)
	If Tit = "" Then Tit = YSvoid.ChannelName
	If Not YSvoid.ChannelTrue Then
		Call Format_Redirect(YSvoid.Web_Dir&"Common/Page_Load.asp?该频道不存在或暂时维护中！")
		Exit Sub
	End If
	'--- 如果频道未禁用并且有模板数据便载入
	SkinTrue = YSvoid.Channel_TemplateLoad(ChannelID)
	'--- 频道出现异常显示错误信息
	If Not SkinTrue Then
		Call Format_Redirect(YSvoid.Web_Dir&"Common/Page_Load.asp?栏目（"&YSvoid.ChannelName&"）模版未配置或出现异常！")
		Exit Sub
	End If
	'--- 载入频道图片大小设置
	Pic_W = Int(YSvoid.ChannelPicWidth)
	Pic_H = Int(YSvoid.ChannelPicHeight)
	'--- 载入专题
	If YSvoid.SpecialTrue Then
		YSvoid.Name = "Channel_Special"&ChannelID
		If YSvoid.Cache_Chk() Then	Call YSvoid.Special_KeyWord_Cast_ToCache(ChannelID,0)
		YSvoid.ChannelSpecial = YSvoid.Value
		If Not IsArray(YSvoid.ChannelSpecial) Then YSvoid.SpecialTrue = False
	End If
	'--- 载入热门关键字
	If YSvoid.KeyWordTrue Then
		YSvoid.Name = "Channel_KeyWord"&ChannelID
		If YSvoid.Cache_Chk() Then	Call YSvoid.Special_KeyWord_Cast_ToCache(ChannelID,1)
		YSvoid.ChannelKeyWord = YSvoid.Value
		If Not IsArray(YSvoid.ChannelKeyWord) Then YSvoid.KeyWordTrue = False
	End If
    '--- 载入公告
	If YSvoid.CastTrue Then
		YSvoid.Name = "Channel_Cast"&ChannelID
		If YSvoid.Cache_Chk() Then	Call YSvoid.Special_KeyWord_Cast_ToCache(ChannelID,2)
		YSvoid.ChannelCast = YSvoid.Value
		If Not IsArray(YSvoid.ChannelCast) Then YSvoid.CastTrue = False
	End If
	'---- 前台判断结束
End Sub

'================================================
'作    用：网站头部功能模块,设置页面头部信息及显示形式,浏览权限等
'参    数：
'	   var1 : 可选三个参数：0，1，2
'               0，表示不要登陆即可进行浏览该页面；
'               1，表示在后台设置要登陆浏览时，表示有权限用户（请查看下面的说明）登陆后才可浏览该页面；
'               2，表示强制有权限用户登陆后才可浏览该项页面；
'	   var2 : 可选两个参数：0，1
'               0，表示不论是否锁定用户可浏览该页面；
'               1，表示已被锁定的用户不可以浏览该页面；
'	   var3 : 可选三个参数：0，1，2，3，4，5
'               0，表示以左小右大形式的方式显示页面，输出头部和底部信息，需配合“call web_center(0)、                                                    call web_end(0,0)”一起使用；
'               1，此参数值保留；
'               2，表示以满屏方式显示，输出头部和底部信息，不分左右，配合“call
'                  web_end(0,0)”一起使用，当设置此值时该页面请勿再使用“call web_center(0)”；
'               3，此参数值保留；
'               4，表示只输出页头JS，不包含头部和底部信息，配合“call
'                  web_end(0,1)”一起使用，当设置此值时该页面请勿再使用“call web_center(0)”；
'               5，表示只进行权限处理，不输出头部和底部信息，配合“call
'                  web_end(0,1)”一起使用，当设置此值时该页面请勿再使用“call web_center(0)”；
'	   var4 : 此参数值保留
'	   var5 : 此参数值保留
'================================================
Sub Web_Head(ByVal var1,ByVal var2,ByVal var3,ByVal var4,ByVal var5)
	Dim UserTit				'---用户所在栏目
	Dim UserExitim			'---用户过期更新时间
	Dim UserIsOnline		'---判断用户是否在线
	Call Web_Head_Tit()
	'Call YSvoid.YSvoid_InfoInit()

	Dim ttt,User_Info,ok_msg,power
	stamp_num=YSvoid.Web_Mode
	'--- 下面为判断用户信息部分
	If Cookies_True Then
		login_password=YSvoid.Code_Encrypt(login_password)
		If ChkName(login_username)=False Or ChkPassWord(login_password)=False Then
			Call Cookies_Clear(0)
		Else
			Dim NeedToUpdate,NeedTrue
			NeedTrue=False
			NeedToUpdate=","&YSvoid.Cache_Get("NeedToUpdate")&","
			If InStr(NeedToUpdate,","&login_username&",")>0 Then
				NeedTrue=True
			End If
			User_Info=YSvoid.Session_Get(login_username)
			If Not IsArray(User_Info) Or NeedTrue=True Then
				Call YSvoid.NeedUpdateList(login_username,0)
				SQL="select top 1 id,username,power,popedom,emoney,integral,last_tim,face,face_w,face_h,estate,abate,otim,bbs_counter,onlinetime,login_msg,email,whe,phone,strength,expression,iask_counter,iask_helpnum,iask_reply from YSvoid_UserData where hidden=1 and username='"&login_username&"' and password='"&login_password&"'"
				Call YSvoid.Exe_Conn(Rs,SQL,3)
				If Rs.Eof Then
					Call Cookies_Clear(0)
				Else
				  'Rs("last_tim")=YSvoid.now_time
					Rs.Update
					User_Info=Rs.GetString(,1, "|||","","")
					'--- (0-12) + 3
					'--- 0:用户ID、1:用户名、2:等级、3:权限、4:币、5:积分、6:最后登陆时间、7:头像、8:头像宽、9:头像高
					'--- 10:计时、11:过期、12:过期时间、13:发贴数、14:用户在线时间、15:短信数、16:email
					'--- 17:所在地、18:电话、19：判断用户体力、20：判断用户经验、21:爱问发布数、
					'--- 22:爱问被采纳数、23:爱问被采纳数
					'--- 24:所在栏目、25:最后登陆时间、26:判断用户是否在线
					User_Info=User_Info & "|||" & tit & "|||" & YSvoid.now_time & "|||" & YSvoid.IsOnline(login_username,0)
					User_Info=Split(User_Info,"|||")
					Call YSvoid.Session_Set(login_username,User_Info)
				End If
				Rs.Close
			End If
			 If Cint(User_Info(10))=1 and YSvoid.Time_Type(User_Info(12),4)<YSvoid.Time_Type(YSvoid.now_time,4) or Cint(User_Info(11))=true then
             sql="update YSvoid_UserData set estate=0,abate=0,power='"&format_powers(0)&"' where username='"&User_Info(1)&"'"
		     call YSvoid.exec(sql,0)
		     ok_msg="\n\n您也同时降级为 "&format_power(format_powers(0),1)&"!"
             Response.Write YSvoid.js_put("alert(""您的计时"&format_power(User_Info(2),1)&"用户帐号已过期！"&ok_msg&""");location.href='?';",1)
			 else
			login_id=User_Info(0)
			login_username=User_Info(1)
			login_mode=User_Info(2)
			login_popedom=User_Info(3)
			login_emoney=User_Info(4)    'Cint(User_Info(4))
			login_integral=User_Info(5)  'Cint(User_Info(5))
			login_lasttime=User_Info(6)
			login_faces=User_Info(7)&"|"&User_Info(8)&"|"&User_Info(9)
			login_is_estate=Cint(User_Info(10))
			login_is_abate=Cint(User_Info(11))
			login_bbs_counter=Cint(User_Info(13))
			login_is_otim=User_Info(12)
			User_OnlineTime=User_Info(14)
			login_msg=User_Info(15)
			login_email=User_Info(16)
			login_whe=User_Info(17)
			login_phone=User_Info(18)
			login_strength=User_Info(19)
            login_expression=User_Info(20)
			login_iask_counter=User_Info(21)
			login_iask_helpnum=User_Info(22)
			login_iask_reply=User_Info(23)
			If Int(login_is_abate)>0 Then
			If Int(DateDiff("d",YSvoid.Time_Type(YSvoid.now_time,4),User_Info(12)))<1 Then login_is_abate=0
			End If
			UserTit=User_Info(24)
			UserExitim=User_Info(25)
			UserIsOnline=User_Info(26)
			End If
			Erase User_Info
			login_modep=Format_Power(login_mode,2)
			YSvoid.Cookies_Set "login_id",login_id
			YSvoid.Cookies_Set "login_mode",login_mode
			YSvoid.Cookies_Set "login_password",YSvoid.Code_Encrypt(login_password)
			Is_Login_true=True
		End If
		If Is_Refur_True=True Then Call time_load()

    '--------------------------------用户在线时间识别----------------------------------------------------------
		today_onlinetime=DateDiff("n",login_lasttime,YSvoid.now_time)
    If Is_Login_True And today_onlinetime>5 Then
      Call YSvoid.exec("update YSvoid_UserData set last_tim='"&YSvoid.now_time&"' where username='"&login_username&"'",0)
      If today_onlinetime<30 Then
        Call YSvoid.exec("update YSvoid_UserData set onlinetime=onlinetime+"&today_onlinetime&" where username='"&login_username&"'",0)
        Call YSvoid.exec("update YSvoid_UserData set strength=strength+"&today_onlinetime&" where username='"&login_username&"'",0)
        Call YSvoid.exec("update YSvoid_UserData set expression=expression+"&today_onlinetime&" where username='"&login_username&"'",0)
      End If
    End If
		'---------------------------------完整型,用户和游客都进行识别,这里识别游客----------------------------------
		If Int(stamp_num)>1 And Not Is_Login_True Then
			ttt=YSvoid.Cookies_Get("guest_name")
			ttt=YSvoid.Code_Html(ttt,1,0)
			If ttt="" Then
				ttt="游客"&Session.SessionID
				Call YSvoid.Cookies_Set("guest_name",ttt)
			End If
			User_Info=YSvoid.Session_Get("guest")
			If Not IsArray(User_Info) Then
				User_Info=ttt & "|||" & tit & "|||" & YSvoid.now_time & "|||" & YSvoid.IsOnline(ttt,0)
				User_Info=Split(User_Info,"|||")
				Call YSvoid.Session_Set("guest",User_Info)
			End If
			UserTit=User_Info(1)
			UserExitim=User_Info(2)
			UserIsOnline=User_Info(3)
			If UserIsOnline Then
				If datediff("s",UserExitim,YSvoid.now_time)>YSvoid.Format_Mid_Num(15)*60 then
					If YSvoid.Format_Mid_Num(9)=1 Then
						YSvoid.total_counter=YSvoid.total_counter+1
					End If
				End If
				Call YSvoid.Exec("update YSvoid_UserLogin set l_where='"&left(tit,50)&"',l_tim_end='"&YSvoid.now_time&"',l_power='other',l_forum_id="&YSvoid.Forum_ID&",l_login_id="&login_id&" where l_username='"&ttt&"'",0)
				User_Info(2)=YSvoid.now_time
			Else
				Call YSvoid.Exec("insert into YSvoid_UserLogin(l_username,l_type,l_where,l_tim_login,l_tim_end,l_ip,l_sys,l_power,l_forum_id,l_login_id) values('"&ttt&"',0,'"&left(tit,50)&"','"&YSvoid.now_time&"','"&YSvoid.now_time&"','"&user_ip&"','"&user_sys&"','other',"&YSvoid.Forum_ID&",0)",0)
				YSvoid.total_counter=YSvoid.total_counter+1
				'更新在线人数
				YSvoid.YSvoid_Online=YSvoid.YSvoid_Online+1
				YSvoid.Cache_Set "YSvoid_Online",YSvoid.YSvoid_Online
				'结束
				User_Info(2)=YSvoid.now_time
				User_Info(3)=True
			End If
			Call YSvoid.Session_Set("guest",User_Info)
			Erase User_Info
		End If
		'---------------------------------功能型,用户识别-------------------------------------
		If Int(stamp_num)>0 And Is_Login_True then
			User_Info=YSvoid.Session_Get(login_username)
			'--- 超过过期时间则更新数据库
			If UserIsOnline Then
				If DateDiff("s",UserExitim,YSvoid.now_time)>YSvoid.Format_Mid_Num(15)*60 Then
					If YSvoid.Format_Mid_Num(9)=1 Then
						YSvoid.total_counter=YSvoid.total_counter+1
					End If
				End If
				Call YSvoid.Exec("update YSvoid_UserLogin set l_where='"&left(tit,50)&"',l_tim_end='"&YSvoid.now_time&"',l_power='"&login_mode&"',l_forum_id="&YSvoid.Forum_ID&",l_login_id="&login_id&" where l_username='"&login_username&"'",0)
				User_Info(25)=YSvoid.now_time
			Else
				Call YSvoid.Exec("insert into YSvoid_UserLogin(l_username,l_type,l_where,l_tim_login,l_tim_end,l_ip,l_sys,l_power,l_forum_id,l_login_id) values('"&login_username&"',1,'"&left(tit,50)&"','"&YSvoid.now_time&"','"&YSvoid.now_time&"','"&user_ip&"','"&user_sys&"','"&login_mode&"',"&YSvoid.Forum_ID&","&login_id&")",0)
				YSvoid.total_counter=YSvoid.total_counter+1
				'更新在线人数
				YSvoid.YSvoid_Online=YSvoid.YSvoid_Online+1
				YSvoid.Cache_Set "YSvoid_Online",YSvoid.YSvoid_Online
				YSvoid.YSvoid_UserOnline=YSvoid.YSvoid_UserOnline+1
				YSvoid.Cache_Set "YSvoid_UserOnline",YSvoid.YSvoid_UserOnline
				'结束
				User_Info(25)=YSvoid.now_time
				User_Info(26)=True
			End If
			Call YSvoid.Session_Set(login_username,User_Info)
			'--- 删除用户Session
			Call YSvoid.Session_Del(login_username)
		End If
	End If
	'--- 删除过期时间用户，更新在线人数
	YSvoid.UpdateOnline
	online_num=YSvoid.YSvoid_Online
	UserOnline_Num=YSvoid.YSvoid_UserOnline

	If Cint(online_num)>Cint(YSvoid.YSvoid_Info(3,0)) Then
		Call YSvoid.ReloadInfo(online_num,3)
		Call YSvoid.ReloadInfo(YSvoid.now_time,4)
		Call YSvoid.Exec("update YSvoid_Configs set max_online="&online_num&",max_tim='"&YSvoid.now_time&"'",0)
	End If
	If Int(YSvoid.total_counter)>Int(YSvoid.YSvoid_Info(2,0)) Then
		Call YSvoid.ReloadInfo(YSvoid.total_counter,2)
		Call YSvoid.Exec("update YSvoid_Configs set counter="&YSvoid.total_counter,0)
	End If

	'--- 检查页面权限
	If Page_Power<>"" And Instr("."&Page_Power&".","."&login_modep&".")<1 Then Call Web_Error("power")

	'-- var1
	'-- 1：需登陆才能浏览信息时
	'-- 2：用户未登陆时
	Select Case var1
	Case 1
		If Int(YSvoid.Format_Mid_Num(8))=1 And Not Is_Login_True Then Call My_Login()
	Case 2
		If Not Is_Login_True Then Call My_Login()
	End Select
	'--- var2
	'--- 1：检查用户是否被锁定，是否有浏览时间限制
	If var2=1 Then
		If Time_Lock=True Then Call Web_Error("time_lock")
		If login_mode<>"" Then
			If popedom_P(111)=1 Then Call Web_Error("locked")
		End If
	End If
	'--- var3
	'--- 5：只显示头部
	If var3=5 Then Exit Sub
	'--- var3
	'--- 4：只显示头部与BODY部分
	If var3=4 Then Exit Sub
	YSvoid.HtmlWhole 1
	YSvoid.HtmlRcod "menu",YSvoid_Menu(tit)
	YSvoid.HtmlRcod "today_week",FormatDateTime(YSvoid.today_time,1)&"&nbsp;"&WeekDayName(WeekDay(YSvoid.today_time))
	If YSvoidTit="" Then YSvoidTit=tit
	If tit_fir="" Then
		YSvoid.HtmlRcod "tit_now","<font title='"&YSvoid.code_html(YSvoidTit,1,0)&"'>"&tit&"</font>"
	Else
		If Index_Url <> "" Then
			YSvoid.HtmlRcod "tit_now","<a href='"&index_url&".asp' class=h_menu>"&tit_fir&"</a>&nbsp;→&nbsp;<font class=h_menu>"&YSvoidTit&"</font>"
		Else
			YSvoid.HtmlRcod "tit_now","<a href='./' class=h_menu>"&tit_fir&"</a>&nbsp;→&nbsp;<font class=h_menu>"&YSvoidTit&"</font>"
		End If
	End If
	YSvoid.HtmlView(0)
	Call Web_Left(var3)
End Sub

'================================================
'作    用：网站底部信息功能模块
'参    数：
'	   var1 : 此参数值保留
'	   var2 : 判断是否显示底部信息
'================================================
Sub Web_End(ByVal var1,ByVal var2)
	If var2=0 Then
		Response.Write web_right
		YSvoid.HtmlWhole 4
		YSvoid.HtmlRcod "exec_tim",YSvoid.Exec_Tim
		YSvoid.HtmlRcod "exec_num",YSvoid.num_rs
		YSvoid.HtmlRcod "online_num",online_num
		YSvoid.HtmlRcod "cache_num",YSvoid.Cache_Num
		YSvoid.HtmlView(0)
		If int(login_msg)>0 And Action<>"msg_view" Then
			Dim msg_id,msg_name,msg_topic
			SQL="select top 1 id,send_u,topic from YSvoid_UserMail where accept_u='"&login_username&"' and types=1 and isread=0"
			Set Rs=YSvoid.Exec(SQL,1)
			if not rs.eof then
			msg_id=Rs(0)
			msg_name=Rs(1)
			msg_topic=Rs(2)
			end if
			Rs.Close
			YSvoid.HtmlMain "login_msg"
			YSvoid.HtmlRcod "login_msg",login_msg
			YSvoid.HtmlRcod "msg_id",msg_id
			YSvoid.HtmlRcod "dir_skin",YSvoid.web_dir_skin
			YSvoid.HtmlRcod "msg_name",msg_name
			YSvoid.HtmlRcod "msg_topic",msg_topic
			YSvoid.HtmlView(0)
		End If
	End If
	If var2<5 Then Call Web_Ender()
	Call Web_Clear(1)
End Sub

'//* 网站底部(结束部分) *//
Private Function Web_Ender()
	YSvoid.HtmlWhole 11
	YSvoid.HtmlView(0)
End Function

'//* 格式化网站左边部分 *//
Private Function Web_Left(ByVal lt)
	Dim tNum
	Select Case Cint(lt)
	Case 1:tNum=6
	Case 2:tNum=7
	Case 3:tNum=8
	Case Else:tNum=5
	End Select
	YSvoid.HtmlWhole tNum
	YSvoid.HtmlView(0)
End Function

'//* 格式化网站中间部分 *//
Public Function Web_Center(ByVal ct)
	YSvoid.HtmlWhole 9
	YSvoid.HtmlView(0)
End Function

'//* 格式化网站右边部分 *//
Private Function Web_Right()
	YSvoid.HtmlWhole 10
	YSvoid.HtmlView(0)
End Function

'//* 格式化表格换行 *//
Public Function Web_Branch(ByVal jst)
	Call Web_Right()
	Call Web_Left(2)
	Response.Buffer=True
End Function

'//* 格式化表格 *//
Public Function Format_Add_Frame(ByVal vBody)
	YSvoid.HtmlMain "format_add_frame"
	YSvoid.HtmlRcod "body",vBody
	format_add_frame=YSvoid.HtmlGet(0)
End Function

'//* 格式化频道搜索 *//
Public Function Format_nSort_Sea(ByVal vTit)
	YSvoid.HtmlMain "format_nsort_sea"
	YSvoid.HtmlRcod "stit",vTit
	format_nsort_sea=YSvoid.HtmlGet(0)
End Function

'//* 等级设置 *//
Public Function popedom_p(ByVal oNum)
	If Cint(oNum)>Len(login_popedom) Or Len(login_popedom)<>120 Then
		popedom_p=0
		Exit Function
	End If
	popedom_p=Int(Mid(login_popedom,oNum,1))
End Function

'================================================
'作    用：弹出对话框
'参    数：
'	   msg_type : 判断弹出方式
'	   msg_var  : 提示内容
'      msg_url  : 跳转路径
'================================================
Public Sub YSvoid_Msg(ByVal mType,ByVal mVar,ByVal mUrl)
	Call Cookies_Load()
	If mType=0 Then
		Response.Write "alert("""&mVar&""");location.href="""&mUrl&""";"
		Exit Sub
	End If
	Response.Write YSvoid.Js_Put("alert("""&mVar&""");location.href="""&mUrl&""";",1)
End Sub

'//* 设置过期时间到 Cookies *//
Public Sub Cookies_Load()
	YSvoid.Cookies_Set "time_load",dateadd("s",-5,YSvoid.now_time)
End Sub

'//* 清除 Cookies *//
Public Sub Cookies_Clear(ByVal tl)
	If tl=1 Then Call Cookies_Load()
	YSvoid.Cookies_Set "login_username",""
	YSvoid.Cookies_Set "login_password",""
	YSvoid.Cookies_Set "login_id",""
	YSvoid.Cookies_Set "login_mode",""
	YSvoid.Cookies_Set "iscookies",""
	login_mode=""
	login_username=""
	login_password=""
End Sub

'//* 未登陆处理信息 *//
Private Sub My_Login()
	If Is_Login_True Then Exit Sub
	Call Cookies_Clear(1)
	Call Format_Redirect(YSvoid.Web_Dir&"User/Login.asp?old_url="&YSvoid.BrowseUrl)
End Sub

'//* 取频道名称 *//
Public Function Format_Menu(ByVal mVar)
	Dim i, TmpStr
	TmpStr=YSvoid.Cache_Get("web_menu")
	Format_Menu=""
	For i=0 To Ubound(TmpStr,2)
		If mVar=TmpStr(1,i) Then
			Format_Menu=TmpStr(2,i)
			Exit For
		End If
	Next
End Function

'//* 页面跳转 *//
Public Sub Format_Redirect(ByVal rUrl)
	Call Web_Clear(0)
	Response.Clear
	Response.Redirect rUrl
	Response.End
End Sub

'//* 处理网站关闭信息 *//
Private Sub YSvoid_IsOpen()
	If Int(YSvoid.Web_Dim(0)) = 0 Then
		Response.Clear
		Tit="暂停使用"
		Call Web_Head_Tit()
		Response.Write YSvoid.Web_Closer
		Call Web_Ender()
		Call Web_Clear(1)
	End If
End Sub

Private Sub YSvoid_TimOpen()
dim Tim_Open,Now_Time
Tim_Open = YSvoid.Web_Dim(7)
Now_Time=hour(now())
if Int(Mid(Tim_Open,Now_Time+1,1))=1 Then
		Response.Clear
		Tit="暂停使用"
		Call Web_Head_Tit()
		Response.Write YSvoid.Web_Closer
		Call Web_Ender()
		Call Web_Clear(1)
	End If
End Sub

'//* 处理用户IP锁定信息 *//
Private Sub web_ip_shield()
	Dim tmp
	tmp=True
	If YSvoid.Is_Null(YSvoid.Web_LockIP) = "" Or user_ip="" Then Exit Sub
	Dim sdim,snum
	sdim=Split(YSvoid.Web_LockIP,"|")
	snum=Ubound(sdim)
	For i=0 To snum
		If Instr("$"&user_ip&".","$"&sdim(i)&".")>0 Then
			Erase sdim
			Response.Clear
			Call Web_Head_Tit()
			YSvoid.HtmlMain "no_ip"
			YSvoid.HtmlView(0)
			Call Web_Clear(1)
		End If
	Next
	Erase sdim
End Sub

'//* 网站错误信息 *//
Public Sub Web_Error(ByVal eType)
	Response.Clear
	YSvoid.Cookies_Set "time_load",DateAdd("s",-5,YSvoid.now_time)
	YSvoid.Cookies_Set "old_url",Request.ServerVariables("http_referer")
	YSvoid.Cookies_Set "error_action",eType
	Call Format_Redirect(YSvoid.Web_Dir&"Common/Error.asp")
End Sub

'//* 用户登陆通用模块 *//
Public Sub User_Login()
	If Not Is_Login_True Then
		YSvoid.HtmlMain "left_login_false"
	    YSvoid.HtmlRcod "valcode",YSvoid.GetCode()
		YSvoid.HtmlView(1)
		Exit Sub
	End If
	Dim msg_var
	If int(login_msg)<1 Then
		msg_var="<font class=gray>"&login_msg&"</font>"
	Else
		msg_var="<font class=red>"&login_msg&"</font>"
	End If
	YSvoid.HtmlMain "left_login_true"
	YSvoid.HtmlRcod "login_username",login_username
	YSvoid.HtmlRcod "login_mode",Format_Power(login_mode,1)
	YSvoid.HtmlRcod "login_emoney",login_emoney
	YSvoid.HtmlRcod "login_integral",login_integral
	YSvoid.HtmlRcod "login_strength",login_strength
	YSvoid.HtmlRcod "login_expression",login_expression
	YSvoid.HtmlRcod "login_msg",msg_var
	YSvoid.HtmlView(1)
End Sub

'================================================
'作    用：网站刷新时间相关
'================================================
Public Sub Time_Load()
	Dim tim_s,var_s
	tim_s=Int(YSvoid.Format_Mid_Num(16))
	If Int(tim_s)<1 Then Exit Sub
	If Login_Mode=Format_Power2(1,1) Then Exit Sub
	var_s=YSvoid.Cookies_Get("time_load")
	If IsDate(var_s) Then
		If Int(DateDiff("s",var_s,YSvoid.now_time))<=Int(tim_s) Then
			Call Web_Error("time_load")
			Exit Sub
		End If
	End If
	YSvoid.Cookies_Set "time_load",YSvoid.now_time
End Sub

'//* 发布信息时间限制 *//
Private Function Time_Lock()
	Dim TmpStr,Tim1
	Time_Lock=False
	TmpStr=YSvoid.Web_Dim(7)
	If Len(TmpStr)<>24 Or Not YSvoid.Is_Int(TmpStr) Then Exit Function
	Tim1=Int(Hour(YSvoid.Now_Time))
	if Tim1=24 Then Tim1=0
	If Int(Mid(TmpStr,Tim1+1,1))=1 Then Time_Lock=True
End Function

'//* 提交信息安全设置 *//
Public Function Frm_Chk()
	Frm_Chk=False
	If YSvoid.Code_Form("chk")="yes" Then
		Frm_Chk=YSvoid.Post_Chk()
	End If
End Function

'================================================
'作    用：图片路径处理
'参    数：
'	   pvar : 判别图片路径
'	   pt   : 判别图片大小
'================================================
Public Function Pic_Url(ByVal pVar,ByVal pt)
	Dim temp1,pic_var
	pic_var=pVar
	If pic_var="" Then
		pic_var=YSvoid.Web_Dir&YSvoid.web_upload&"/no_pic.gif"
	End If

    if isNumeric(pt)=false and instr(pt,"*")<>0 then
       pt=replace(pt,"*"," height=" )
       temp1="<img border=0 width="+pt+" src='" + pic_var + "'>"
    else
	  Select Case Int(pt)
	  Case 1
		temp1="<img border=0 width="&pic_w&" height="&pic_h&" src='" & pic_var & "'>"
	  Case 2
		temp1="<img border=0 width="&(pic_w*1.5)&" height="&(pic_h*1.5)&" src='" & pic_var & "'>"
	  Case 3
		temp1="<img border=0 src='" & pic_var & "'>"
	  Case 4
		temp1="<img border=0 width="&(pic_w*0.7)&" height="&(pic_h*0.7)&" src='" & pic_var & "'>"
	  Case 5
		temp1="<img border=0 width="&(pic_w*0.5)&" height="&(pic_h*0.5)&" src='" & pic_var & "'>"
      Case 6
		temp1="<img border=0 width="&(pic_w*1.2)&" height="&(pic_h*0.9)&" src='" & pic_var & "'>"
	  Case 7
		temp1="<img border=0 width="&(pic_w*0.8)&" height="&(pic_h*0.8)&" src='" & pic_var & "'>"
	  Case 8
		temp1="<img border=0 width="&(pic_w*1.5)&" height="&(pic_h*1.5)&" src='" & pic_var & "'>"
      Case 9
		temp1="<img border=0 width="&(pic_w*0.91)&" height="&(pic_h*1.58)&" src='" & pic_var & "'>"
	  Case 10
		temp1="<img border=0 width="&(pic_w*1.7)&" height="&(pic_h*3)&" src='" & pic_var & "'>"
	  Case 11
		temp1="<img border=0 width="&(pic_w*0.9)&" height="&(pic_h*0.7)&" src='" & pic_var & "'>"
	  Case Else
		temp1=pic_var
	  End Select
	End If
	Pic_Url=temp1
End Function

'================================================
'作    用：图片边框处理
'参    数：
'	   pvar : 判别图片路径
'	   pt   : 判别图片大小
'	   purl : 判断图片路径
'================================================
Public Function Pic_Fk(ByVal pVar,ByVal pt,ByVal pUrl)
	dim burl,picurl
	burl=YSvoid.Web_Dir&"Images/Fk/"
	picurl=Pic_Url(pVar,pt)
	If pUrl<>"" And YSvoid.Is_Null(pUrl)<>"" Then picurl="<a href='"&purl&"' target=_blank>"&picurl&"</a>"
	YSvoid.HtmlMain "pic_fk"
	YSvoid.HtmlRcod "burl",burl
	YSvoid.HtmlRcod "picurl",picurl
	Pic_Fk=YSvoid.HtmlGet(0)
End Function

'================================================
'作    用：标题模块
'参    数：
'	   b_jt         : 标题前小图片
'	   b_username   : 发布人
'	   b_topic      : 标题
'	   b_c_num      : 标题显示字数
'	   b_url        : 连接地址
'	   b_tim        : 发布时间
'	   b_counter    : 浏览的次数
'	   b_ispic      : 缩略图
'	   b_tit        : 标签
'	   b_count      : 回复、浏览等
'	   b_tim_type   : 发布时间
'================================================
Public Function Format_Topic_Type(b_jt,b_info,b_username,b_topic,b_c_num,b_url,b_tim,b_color,b_ispic,b_tit,b_tim_type,b_pic)
	Dim n_img,tim_type,n_c_num,temp1
	If b_ispic=true Then
	n_c_num=b_c_num
		n_img="&nbsp;<img src='"&YSvoid.web_dir_skin&"small/img.gif' border=0 Title='<img src="&b_pic&" border=0 height=80>'>"
		n_c_num=n_c_num-2
	End If
	If Int(b_tim_type)>0 Then
		If Int(b_tim_type)=2 Then tim_type=tim_type&"</td><td align=right>"
		tim_type=tim_type&"&nbsp;<font class=tims>"&YSvoid.time_type(b_tim,7)&"</font>"
	End If
	YSvoid.HtmlMain "topic"
	YSvoid.HtmlRcod "t_height",m_hei
	YSvoid.HtmlRcod "t_jt",b_jt
	YSvoid.HtmlRcod "t_info",b_info
	YSvoid.HtmlRcod "t_url",b_url
	YSvoid.HtmlRcod "t_tit",b_tit
	YSvoid.HtmlRcod "t_topic",YSvoid.code_html(b_topic,1,b_c_num)
	YSvoid.HtmlRcod "t_topics",YSvoid.code_html(b_topic,1,0)
	YSvoid.HtmlRcod "t_username",b_username
	YSvoid.HtmlRcod "t_color",b_color
	YSvoid.HtmlRcod "t_tim",b_tim
	YSvoid.HtmlRcod "t_img",n_img
	YSvoid.HtmlRcod "t_tim_type",tim_type
	Format_Topic_Type=YSvoid.HtmlGet(0)
End Function

'================================================
'作    用：图片模块
'参    数：
'	   b_pic        : 图片地址
'	   b_topic      : 图片标题
'	   b_c_num      : 标题显示字数
'	   b_url        : 图片连接地址
'================================================
Public Function Format_Pic_Type(b_pic,b_topic,b_c_num,b_url)
	Dim TmpPic
	TmpPic = Pic_Fk(b_pic,1,b_url)
	YSvoid.HtmlMain "format_pic"
	YSvoid.HtmlRcod "pic_fk",TmpPic
	YSvoid.HtmlRcod "t_url",b_url
	YSvoid.HtmlRcod "t_topics",YSvoid.Code_Html(b_topic,1,0)
	YSvoid.HtmlRcod "t_topic",YSvoid.Code_Html(b_topic,1,b_c_num)
	Format_Pic_Type=YSvoid.HtmlGet(0)
End Function

'//* 网站版权信息 *//
Public Function YSvoid_Copyright()
	YSvoid.HtmlMain "copyright"
	YSvoid.HtmlRcod "closer",YSvoid.ModGet("closer")
	YSvoid.HtmlView(0)
End Function

'//* 网站搜索 *//
sub Serach_Dhang()
    Dim tmp_Channel,tmp_types,iListStr,Lii,tmp_txt
	iListStr = YSvoid.Channel_Menu
	tmp_Channel=vbcrlf&"<select name=sea_type>" 
	For Lii =0 To Ubound(iListStr,2)
	tmp_types=lcase(iListStr(12,Lii))
    tmp_txt=iListStr(13,Lii)
    tmp_Channel=tmp_Channel&vbcrlf&"<option value='"&tmp_types&"'>"&tmp_txt&"</option>" 
    Next
	erase iListStr
	tmp_Channel=tmp_Channel&vbcrlf&"</select>"
	YSvoid.HtmlMain "search_dhang"
	YSvoid.HtmlRcod "Channel",tmp_Channel
	YSvoid.HtmlView(1)
end sub

'//* 网站菜单 *//
Public Function YSvoid_Menu(tit)
	Dim temp1,tmp1,TmpStr
	Dim Cii, oChannelType
	Dim oChannelName, oChannelPicUrl, oChannelRemark, oChannelColor, oChannelBold, oChannelOpenType, oChannelLinkUrl
	TmpStr = YSvoid.Channel_Menu
	For Cii = 0 To Ubound(TmpStr,2)
		oChannelType = Int(TmpStr(8,Cii))
		If TmpStr(17,Cii) = True Or oChannelType = 2 Then
			If oChannelType = 2 Then
				oChannelLinkUrl = TmpStr(9,Cii)
			Else
				oChannelLinkUrl = YSvoid.Web_Dir & TmpStr(12,Cii) & "/"   'LCase(TmpStr(12,Cii))转换小写
			End If
			temp1="<a href='"&oChannelLinkUrl&"'"
			If TmpStr(3,Cii) <> "" Then temp1 = temp1 & " title='"&YSvoid.Code_Html(TmpStr(3,Cii),1,0)&"'"
			If TmpStr(6,Cii) = True Then temp1 = temp1 & " target=_blank"
			temp1 = temp1 & "><span>"
			oChannelName = YSvoid.Code_Html(TmpStr(1,Cii),1,0)
			If TmpStr(2,Cii) <> "" Then
				oChannelName = "<img src='"&TmpStr(2,Cii)&"'>"
			Else
				If TmpStr(5,Cii) = True Then oChannelName = "<b>" & oChannelName & "</b>"
				If TmpStr(4,Cii) <> "" Then oChannelName = "<font color='#"&TmpStr(4,Cii)&"'>" & oChannelName & "</font>"
			End If
			temp1 = temp1 & oChannelName & "</span></a>"
			tmp1 = tmp1 & VbCrlf & " <li id='menu"&TmpStr(12,Cii)&"'>" & temp1 & " </li>"
		End If
	Next
	Erase TmpStr
	If tit = "首页" Then
	tmp1="<li id='menuHome'><li id='menuHomes'><a href='"&YSvoid.Web_Dir&"'><span>首页</span></a></li></li>" &tmp1 
   Else
   tmp1="<li><li id='menuHome'><a href='"&YSvoid.Web_Dir&"'><span>首页</span></a></li></li>" &tmp1 
   End If
	YSvoid_Menu=tmp1
End Function

%>