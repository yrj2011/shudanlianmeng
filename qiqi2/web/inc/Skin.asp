<!-- #Include File="Config.asp" -->
<!-- #Include File="Functions.asp" -->
<%

Dim m_channel,web_img_m,web_img_h
Dim login_is_abate,login_is_estate,login_is_otim

'// �û�IP��Ϣ
Dim user_ip
	user_ip = YSvoid.Ip_Sys(0)

'// �û�ϵͳ��Ϣ
Dim user_sys
	user_sys = YSvoid.Code_Var(Left(YSvoid.Ip_Sys(1),250))

'// Cookies֧��
Dim Cookies_True
	Cookies_True = True
	If YSvoid.Cookies_Get("cookies_true")<>"yes" Then
		Response.Cookies(YSvoid.web_cookies).Path = YSvoid.Web_Dir
		YSvoid.Cookies_Set "cookies_true","yes"
		Cookies_True = False
	End If

'// Frame ��ʾ/����
Dim iFrame
	iFrame = False

'// ���ñ�������
Dim pic_w,pic_h,max_w,max_h,img_emoney,m_hei
  pic_w = YSvoid.Format_Mid_Num(22)
  If int(YSvoid.Format_Mid_Num(22))<>int(YSvoid.ChannelPicWidth) and int(YSvoid.ChannelPicWidth)>0 Then pic_w = YSvoid.ChannelPicWidth
  pic_h = YSvoid.Format_Mid_Num(23)
  If int(YSvoid.Format_Mid_Num(23))<>int(YSvoid.ChannelPicHeight) and int(YSvoid.ChannelPicHeight)>0 Then pic_h = YSvoid.ChannelPicHeight
  max_w = YSvoid.Format_Mid_Num(24)
  m_hei = YSvoid.Format_Mid_Num(13)
  img_emoney = "<img border=0 src='"&YSvoid.web_dir_skin&"small/emoney.gif' alt='"&YSvoid.web_unit&"' align=absmiddle>"

'// �û��Ƿ��½
Dim Is_Login_true
	Is_Login_true = False

'// ˢ������
Dim Is_Refur_True
	Is_Refur_True = True

'// ��վ��������
Dim Stamp_Num

'// ���������������û���
Dim Online_Num,UserOnline_Num
	Online_Num = 0
	UserOnline_Num = 0

'// ��վ�Ƿ���
Call YSvoid_IsOpen()

'// ��վ����ʱ���Ƿ���
Call YSvoid_TimOpen()

'// �û� IP �Ƿ�����
Call web_ip_shield()

'//* ��վ��ʶ *//
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
'�� ;��վ�ڹ��ģ��
'�� ����
'	aSort���������
'	aType���Ƿ�����
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
		Call Format_Redirect(YSvoid.Web_Dir&"Common/Page_Load.asp?��Ƶ�������ڻ���ʱά���У�")
		Exit Sub
	End If
	'--- ���Ƶ��δ���ò�����ģ�����ݱ�����
	SkinTrue = YSvoid.Channel_TemplateLoad(ChannelID)
	'--- Ƶ�������쳣��ʾ������Ϣ
	If Not SkinTrue Then
		Call Format_Redirect(YSvoid.Web_Dir&"Common/Page_Load.asp?��Ŀ��"&YSvoid.ChannelName&"��ģ��δ���û�����쳣��")
		Exit Sub
	End If
	'--- ����Ƶ��ͼƬ��С����
	Pic_W = Int(YSvoid.ChannelPicWidth)
	Pic_H = Int(YSvoid.ChannelPicHeight)
	'--- ����ר��
	If YSvoid.SpecialTrue Then
		YSvoid.Name = "Channel_Special"&ChannelID
		If YSvoid.Cache_Chk() Then	Call YSvoid.Special_KeyWord_Cast_ToCache(ChannelID,0)
		YSvoid.ChannelSpecial = YSvoid.Value
		If Not IsArray(YSvoid.ChannelSpecial) Then YSvoid.SpecialTrue = False
	End If
	'--- �������Źؼ���
	If YSvoid.KeyWordTrue Then
		YSvoid.Name = "Channel_KeyWord"&ChannelID
		If YSvoid.Cache_Chk() Then	Call YSvoid.Special_KeyWord_Cast_ToCache(ChannelID,1)
		YSvoid.ChannelKeyWord = YSvoid.Value
		If Not IsArray(YSvoid.ChannelKeyWord) Then YSvoid.KeyWordTrue = False
	End If
    '--- ���빫��
	If YSvoid.CastTrue Then
		YSvoid.Name = "Channel_Cast"&ChannelID
		If YSvoid.Cache_Chk() Then	Call YSvoid.Special_KeyWord_Cast_ToCache(ChannelID,2)
		YSvoid.ChannelCast = YSvoid.Value
		If Not IsArray(YSvoid.ChannelCast) Then YSvoid.CastTrue = False
	End If
	'---- ǰ̨�жϽ���
End Sub

'================================================
'��    �ã���վͷ������ģ��,����ҳ��ͷ����Ϣ����ʾ��ʽ,���Ȩ�޵�
'��    ����
'	   var1 : ��ѡ����������0��1��2
'               0����ʾ��Ҫ��½���ɽ��������ҳ�棻
'               1����ʾ�ں�̨����Ҫ��½���ʱ����ʾ��Ȩ���û�����鿴�����˵������½��ſ������ҳ�棻
'               2����ʾǿ����Ȩ���û���½��ſ��������ҳ�棻
'	   var2 : ��ѡ����������0��1
'               0����ʾ�����Ƿ������û��������ҳ�棻
'               1����ʾ�ѱ��������û������������ҳ�棻
'	   var3 : ��ѡ����������0��1��2��3��4��5
'               0����ʾ����С�Ҵ���ʽ�ķ�ʽ��ʾҳ�棬���ͷ���͵ײ���Ϣ������ϡ�call web_center(0)��                                                    call web_end(0,0)��һ��ʹ�ã�
'               1���˲���ֵ������
'               2����ʾ��������ʽ��ʾ�����ͷ���͵ײ���Ϣ���������ң���ϡ�call
'                  web_end(0,0)��һ��ʹ�ã������ô�ֵʱ��ҳ��������ʹ�á�call web_center(0)����
'               3���˲���ֵ������
'               4����ʾֻ���ҳͷJS��������ͷ���͵ײ���Ϣ����ϡ�call
'                  web_end(0,1)��һ��ʹ�ã������ô�ֵʱ��ҳ��������ʹ�á�call web_center(0)����
'               5����ʾֻ����Ȩ�޴��������ͷ���͵ײ���Ϣ����ϡ�call
'                  web_end(0,1)��һ��ʹ�ã������ô�ֵʱ��ҳ��������ʹ�á�call web_center(0)����
'	   var4 : �˲���ֵ����
'	   var5 : �˲���ֵ����
'================================================
Sub Web_Head(ByVal var1,ByVal var2,ByVal var3,ByVal var4,ByVal var5)
	Dim UserTit				'---�û�������Ŀ
	Dim UserExitim			'---�û����ڸ���ʱ��
	Dim UserIsOnline		'---�ж��û��Ƿ�����
	Call Web_Head_Tit()
	'Call YSvoid.YSvoid_InfoInit()

	Dim ttt,User_Info,ok_msg,power
	stamp_num=YSvoid.Web_Mode
	'--- ����Ϊ�ж��û���Ϣ����
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
					'--- 0:�û�ID��1:�û�����2:�ȼ���3:Ȩ�ޡ�4:�ҡ�5:���֡�6:����½ʱ�䡢7:ͷ��8:ͷ���9:ͷ���
					'--- 10:��ʱ��11:���ڡ�12:����ʱ�䡢13:��������14:�û�����ʱ�䡢15:��������16:email
					'--- 17:���ڵء�18:�绰��19���ж��û�������20���ж��û����顢21:���ʷ�������
					'--- 22:���ʱ���������23:���ʱ�������
					'--- 24:������Ŀ��25:����½ʱ�䡢26:�ж��û��Ƿ�����
					User_Info=User_Info & "|||" & tit & "|||" & YSvoid.now_time & "|||" & YSvoid.IsOnline(login_username,0)
					User_Info=Split(User_Info,"|||")
					Call YSvoid.Session_Set(login_username,User_Info)
				End If
				Rs.Close
			End If
			 If Cint(User_Info(10))=1 and YSvoid.Time_Type(User_Info(12),4)<YSvoid.Time_Type(YSvoid.now_time,4) or Cint(User_Info(11))=true then
             sql="update YSvoid_UserData set estate=0,abate=0,power='"&format_powers(0)&"' where username='"&User_Info(1)&"'"
		     call YSvoid.exec(sql,0)
		     ok_msg="\n\n��Ҳͬʱ����Ϊ "&format_power(format_powers(0),1)&"!"
             Response.Write YSvoid.js_put("alert(""���ļ�ʱ"&format_power(User_Info(2),1)&"�û��ʺ��ѹ��ڣ�"&ok_msg&""");location.href='?';",1)
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

    '--------------------------------�û�����ʱ��ʶ��----------------------------------------------------------
		today_onlinetime=DateDiff("n",login_lasttime,YSvoid.now_time)
    If Is_Login_True And today_onlinetime>5 Then
      Call YSvoid.exec("update YSvoid_UserData set last_tim='"&YSvoid.now_time&"' where username='"&login_username&"'",0)
      If today_onlinetime<30 Then
        Call YSvoid.exec("update YSvoid_UserData set onlinetime=onlinetime+"&today_onlinetime&" where username='"&login_username&"'",0)
        Call YSvoid.exec("update YSvoid_UserData set strength=strength+"&today_onlinetime&" where username='"&login_username&"'",0)
        Call YSvoid.exec("update YSvoid_UserData set expression=expression+"&today_onlinetime&" where username='"&login_username&"'",0)
      End If
    End If
		'---------------------------------������,�û����οͶ�����ʶ��,����ʶ���ο�----------------------------------
		If Int(stamp_num)>1 And Not Is_Login_True Then
			ttt=YSvoid.Cookies_Get("guest_name")
			ttt=YSvoid.Code_Html(ttt,1,0)
			If ttt="" Then
				ttt="�ο�"&Session.SessionID
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
				'������������
				YSvoid.YSvoid_Online=YSvoid.YSvoid_Online+1
				YSvoid.Cache_Set "YSvoid_Online",YSvoid.YSvoid_Online
				'����
				User_Info(2)=YSvoid.now_time
				User_Info(3)=True
			End If
			Call YSvoid.Session_Set("guest",User_Info)
			Erase User_Info
		End If
		'---------------------------------������,�û�ʶ��-------------------------------------
		If Int(stamp_num)>0 And Is_Login_True then
			User_Info=YSvoid.Session_Get(login_username)
			'--- ��������ʱ����������ݿ�
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
				'������������
				YSvoid.YSvoid_Online=YSvoid.YSvoid_Online+1
				YSvoid.Cache_Set "YSvoid_Online",YSvoid.YSvoid_Online
				YSvoid.YSvoid_UserOnline=YSvoid.YSvoid_UserOnline+1
				YSvoid.Cache_Set "YSvoid_UserOnline",YSvoid.YSvoid_UserOnline
				'����
				User_Info(25)=YSvoid.now_time
				User_Info(26)=True
			End If
			Call YSvoid.Session_Set(login_username,User_Info)
			'--- ɾ���û�Session
			Call YSvoid.Session_Del(login_username)
		End If
	End If
	'--- ɾ������ʱ���û���������������
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

	'--- ���ҳ��Ȩ��
	If Page_Power<>"" And Instr("."&Page_Power&".","."&login_modep&".")<1 Then Call Web_Error("power")

	'-- var1
	'-- 1�����½���������Ϣʱ
	'-- 2���û�δ��½ʱ
	Select Case var1
	Case 1
		If Int(YSvoid.Format_Mid_Num(8))=1 And Not Is_Login_True Then Call My_Login()
	Case 2
		If Not Is_Login_True Then Call My_Login()
	End Select
	'--- var2
	'--- 1������û��Ƿ��������Ƿ������ʱ������
	If var2=1 Then
		If Time_Lock=True Then Call Web_Error("time_lock")
		If login_mode<>"" Then
			If popedom_P(111)=1 Then Call Web_Error("locked")
		End If
	End If
	'--- var3
	'--- 5��ֻ��ʾͷ��
	If var3=5 Then Exit Sub
	'--- var3
	'--- 4��ֻ��ʾͷ����BODY����
	If var3=4 Then Exit Sub
	YSvoid.HtmlWhole 1
	YSvoid.HtmlRcod "menu",YSvoid_Menu(tit)
	YSvoid.HtmlRcod "today_week",FormatDateTime(YSvoid.today_time,1)&"&nbsp;"&WeekDayName(WeekDay(YSvoid.today_time))
	If YSvoidTit="" Then YSvoidTit=tit
	If tit_fir="" Then
		YSvoid.HtmlRcod "tit_now","<font title='"&YSvoid.code_html(YSvoidTit,1,0)&"'>"&tit&"</font>"
	Else
		If Index_Url <> "" Then
			YSvoid.HtmlRcod "tit_now","<a href='"&index_url&".asp' class=h_menu>"&tit_fir&"</a>&nbsp;��&nbsp;<font class=h_menu>"&YSvoidTit&"</font>"
		Else
			YSvoid.HtmlRcod "tit_now","<a href='./' class=h_menu>"&tit_fir&"</a>&nbsp;��&nbsp;<font class=h_menu>"&YSvoidTit&"</font>"
		End If
	End If
	YSvoid.HtmlView(0)
	Call Web_Left(var3)
End Sub

'================================================
'��    �ã���վ�ײ���Ϣ����ģ��
'��    ����
'	   var1 : �˲���ֵ����
'	   var2 : �ж��Ƿ���ʾ�ײ���Ϣ
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

'//* ��վ�ײ�(��������) *//
Private Function Web_Ender()
	YSvoid.HtmlWhole 11
	YSvoid.HtmlView(0)
End Function

'//* ��ʽ����վ��߲��� *//
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

'//* ��ʽ����վ�м䲿�� *//
Public Function Web_Center(ByVal ct)
	YSvoid.HtmlWhole 9
	YSvoid.HtmlView(0)
End Function

'//* ��ʽ����վ�ұ߲��� *//
Private Function Web_Right()
	YSvoid.HtmlWhole 10
	YSvoid.HtmlView(0)
End Function

'//* ��ʽ������� *//
Public Function Web_Branch(ByVal jst)
	Call Web_Right()
	Call Web_Left(2)
	Response.Buffer=True
End Function

'//* ��ʽ����� *//
Public Function Format_Add_Frame(ByVal vBody)
	YSvoid.HtmlMain "format_add_frame"
	YSvoid.HtmlRcod "body",vBody
	format_add_frame=YSvoid.HtmlGet(0)
End Function

'//* ��ʽ��Ƶ������ *//
Public Function Format_nSort_Sea(ByVal vTit)
	YSvoid.HtmlMain "format_nsort_sea"
	YSvoid.HtmlRcod "stit",vTit
	format_nsort_sea=YSvoid.HtmlGet(0)
End Function

'//* �ȼ����� *//
Public Function popedom_p(ByVal oNum)
	If Cint(oNum)>Len(login_popedom) Or Len(login_popedom)<>120 Then
		popedom_p=0
		Exit Function
	End If
	popedom_p=Int(Mid(login_popedom,oNum,1))
End Function

'================================================
'��    �ã������Ի���
'��    ����
'	   msg_type : �жϵ�����ʽ
'	   msg_var  : ��ʾ����
'      msg_url  : ��ת·��
'================================================
Public Sub YSvoid_Msg(ByVal mType,ByVal mVar,ByVal mUrl)
	Call Cookies_Load()
	If mType=0 Then
		Response.Write "alert("""&mVar&""");location.href="""&mUrl&""";"
		Exit Sub
	End If
	Response.Write YSvoid.Js_Put("alert("""&mVar&""");location.href="""&mUrl&""";",1)
End Sub

'//* ���ù���ʱ�䵽 Cookies *//
Public Sub Cookies_Load()
	YSvoid.Cookies_Set "time_load",dateadd("s",-5,YSvoid.now_time)
End Sub

'//* ��� Cookies *//
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

'//* δ��½������Ϣ *//
Private Sub My_Login()
	If Is_Login_True Then Exit Sub
	Call Cookies_Clear(1)
	Call Format_Redirect(YSvoid.Web_Dir&"User/Login.asp?old_url="&YSvoid.BrowseUrl)
End Sub

'//* ȡƵ������ *//
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

'//* ҳ����ת *//
Public Sub Format_Redirect(ByVal rUrl)
	Call Web_Clear(0)
	Response.Clear
	Response.Redirect rUrl
	Response.End
End Sub

'//* ������վ�ر���Ϣ *//
Private Sub YSvoid_IsOpen()
	If Int(YSvoid.Web_Dim(0)) = 0 Then
		Response.Clear
		Tit="��ͣʹ��"
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
		Tit="��ͣʹ��"
		Call Web_Head_Tit()
		Response.Write YSvoid.Web_Closer
		Call Web_Ender()
		Call Web_Clear(1)
	End If
End Sub

'//* �����û�IP������Ϣ *//
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

'//* ��վ������Ϣ *//
Public Sub Web_Error(ByVal eType)
	Response.Clear
	YSvoid.Cookies_Set "time_load",DateAdd("s",-5,YSvoid.now_time)
	YSvoid.Cookies_Set "old_url",Request.ServerVariables("http_referer")
	YSvoid.Cookies_Set "error_action",eType
	Call Format_Redirect(YSvoid.Web_Dir&"Common/Error.asp")
End Sub

'//* �û���½ͨ��ģ�� *//
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
'��    �ã���վˢ��ʱ�����
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

'//* ������Ϣʱ������ *//
Private Function Time_Lock()
	Dim TmpStr,Tim1
	Time_Lock=False
	TmpStr=YSvoid.Web_Dim(7)
	If Len(TmpStr)<>24 Or Not YSvoid.Is_Int(TmpStr) Then Exit Function
	Tim1=Int(Hour(YSvoid.Now_Time))
	if Tim1=24 Then Tim1=0
	If Int(Mid(TmpStr,Tim1+1,1))=1 Then Time_Lock=True
End Function

'//* �ύ��Ϣ��ȫ���� *//
Public Function Frm_Chk()
	Frm_Chk=False
	If YSvoid.Code_Form("chk")="yes" Then
		Frm_Chk=YSvoid.Post_Chk()
	End If
End Function

'================================================
'��    �ã�ͼƬ·������
'��    ����
'	   pvar : �б�ͼƬ·��
'	   pt   : �б�ͼƬ��С
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
'��    �ã�ͼƬ�߿���
'��    ����
'	   pvar : �б�ͼƬ·��
'	   pt   : �б�ͼƬ��С
'	   purl : �ж�ͼƬ·��
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
'��    �ã�����ģ��
'��    ����
'	   b_jt         : ����ǰСͼƬ
'	   b_username   : ������
'	   b_topic      : ����
'	   b_c_num      : ������ʾ����
'	   b_url        : ���ӵ�ַ
'	   b_tim        : ����ʱ��
'	   b_counter    : ����Ĵ���
'	   b_ispic      : ����ͼ
'	   b_tit        : ��ǩ
'	   b_count      : �ظ��������
'	   b_tim_type   : ����ʱ��
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
'��    �ã�ͼƬģ��
'��    ����
'	   b_pic        : ͼƬ��ַ
'	   b_topic      : ͼƬ����
'	   b_c_num      : ������ʾ����
'	   b_url        : ͼƬ���ӵ�ַ
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

'//* ��վ��Ȩ��Ϣ *//
Public Function YSvoid_Copyright()
	YSvoid.HtmlMain "copyright"
	YSvoid.HtmlRcod "closer",YSvoid.ModGet("closer")
	YSvoid.HtmlView(0)
End Function

'//* ��վ���� *//
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

'//* ��վ�˵� *//
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
				oChannelLinkUrl = YSvoid.Web_Dir & TmpStr(12,Cii) & "/"   'LCase(TmpStr(12,Cii))ת��Сд
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
	If tit = "��ҳ" Then
	tmp1="<li id='menuHome'><li id='menuHomes'><a href='"&YSvoid.Web_Dir&"'><span>��ҳ</span></a></li></li>" &tmp1 
   Else
   tmp1="<li><li id='menuHome'><a href='"&YSvoid.Web_Dir&"'><span>��ҳ</span></a></li></li>" &tmp1 
   End If
	YSvoid_Menu=tmp1
End Function

%>