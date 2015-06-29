<%

'//* 格式化SELECT *//
'--* sName：	Select Name
'--* sView：	显示项目
'--* sValue：	项目值
'--* iValue：	默认选定项目值
Private Function SelectInit(ByVal sName,ByVal sView,ByVal sValue,ByVal iValue)
	Dim Nii
	SelectInit = "<select name="""&sName&""" size=1>"
	For Nii = 0 To Ubound(sView)
		SelectInit = SelectInit & "<option value='"&sValue(Nii)&"'"
		If Cstr(sValue(Nii)) = Cstr(iValue) Then SelectInit = SelectInit & " selected"
		SelectInit = SelectInit & ">"&sView(Nii)&"</option>"
	Next
	SelectInit = SelectInit & "</select>"
End Function

'//* 格式化Input(radio部分) *//
'--* sName：	Select Name
'--* sView：	显示项目
'--* sValue：	项目值
'--* iValue：	默认选定项目值
Private Function InputInit(ByVal sName,ByVal sView,ByVal sValue,ByVal iValue)
	Dim Nii
	For Nii = 0 To Ubound(sView)
		InputInit = InputInit & "<input Type=radio name="""&sName&""" value='"&sValue(Nii)&"'"
		If sValue(Nii) = iValue Then InputInit = InputInit & " checked"
		InputInit = InputInit & ">&nbsp;"&sView(Nii)&"&nbsp;"
	Next
	InputInit = InputInit & "&nbsp;&nbsp;&nbsp;"
End Function

'================================================
'作  用：查看用户信息
'参  数：
'	uuser : 用户名称
'	ut    : 是否弹出窗口显示
'   un    : 是否对用户名称进行截取
'================================================
Public Function Format_User_View(ByVal uuser,ByVal ut,ByVal un)
	If YSvoid.Is_Null(uuser)="" Then
		Format_User_View="<font class=gray>-----</font>"
		Exit Function
	End If
	Dim uname
	uname=uuser
	If un=1 Then uname=YSvoid.Cuted(uuser,10)
	Format_User_View="<a href='"&YSvoid.Web_Dir&"User/View.asp?username="&Server.UrlEncode(uuser)&"' title='查看 "&uuser&" 的详细资料'"
	If ut=1 Then Format_User_View=Format_User_View&" target=_blank"
	Format_User_View=Format_User_View&">"&uname&"</a>"
End Function

'================================================
'作  用：查看用户信息
'参  数：
'	varss : 用户名称
'	ctt   : 是否对用户名称进行截取
'   ct    : 名称显示颜色标记
'================================================
Public Function Format_User_Name(ByVal varss,ByVal ctt,ByVal ct)
	Dim cnum,vars,classt
	cnum=10
	If Int(ctt)=0 Then cnum=0
	vars=varss
	If YSvoid.Is_Null(vars)="" Then
		Format_User_Name="<font class=gray>-----</font>"
		Exit Function
	End If
	If YSvoid.Is_Null(ct)<>"" And ct<>"" Then classt=" class="&ct
	Format_User_Name="<a href='"&YSvoid.Web_Dir&"User/View.asp?username="&Server.UrlEncode(vars)&"' title='查看 "&vars&" 的详细资料' target=_blank"&classt&">"&YSvoid.Code_Html(vars,1,cnum)&"</a>"
End Function

'================================================
'作  用：定义用户图像
'参  数：
'	f_vars: 图像地址
'	f_w   : 图像宽度
'   f_h   : 图像高度
'================================================
Public Function Format_User_Face(ByVal f_vars,ByVal f_w,ByVal f_h)
	If YSvoid.Is_Null(f_vars)="" Then
		Format_User_Face=YSvoid.Web_Dir&"Images/Face/0.gIf"
		Exit Function
	End If
	Format_User_Face="<img src='"&f_vars&"' border=0 width="&f_w&" height="&f_h&">"
End Function

'================================================
'作  用：推荐信息给好友
'参  数：
'	t1: 信息分类
'	t2: 信息主题
'   t3: 信息链接地址
'================================================
Public Function Format_Commend(ByVal t1,ByVal t2,ByVal t3)
	Format_Commend="【<a href='javascript:;' onclick=""javascript:open_win('"&YSvoid.Web_Dir&"Common/Commend.asp?ntit="&Server.UrlEncode(t1)&"&topic="&Server.UrlEncode(t2)&"&url="&Server.UrlEncode(t3)&"','Commend',530,420,'no');"">告诉好友</a>】"
End Function

'================================================
'作  用：举报错误信息与链接
'参  数：
'	t1: 信息分类
'	t2: 信息主题
'   t3: 报错信息编号
'	t4: 报错信息分类
'   t5: 信息链接地址
'================================================
Public Function Commend_Error(ByVal t1,ByVal t2,ByVal t3,ByVal t4)
	Commend_Error="【<a href='javascript:;' onclick=""javascript:open_win('"&YSvoid.Web_Dir&"Common/Err.asp?ChannelID="&t1&"&topic="&Server.UrlEncode(t2)&"&eid="&t3&"&ntit="&Server.UrlEncode(t4)&"','Commend',530,355,'no');"">错误报告</a>】"
End Function

'================================================
'作  用：定义网站图片
'参  数：
'	snum: 图片名称前缀
'================================================
Public Function Img_Small(ByVal sNum)
	Img_Small="<img border=0 src='"&YSvoid.Web_Dir&"Images/Small/"&sNum&".gIf' align=absmiddle>&nbsp;"
End Function

Public Function Img_Skin(ByVal sNum)
	Img_Skin="<img border=0 src='"&YSvoid.web_dir_skin&"Small/"&sNum&".gIf' align=absmiddle>&nbsp;"
End Function

Public Function Img_Alt(ByVal sNum,ByVal sAlt)
	Img_Alt="<img border=0 src='"&YSvoid.web_dir_skin&"Small/"&sNum&".gIf' title='"&sAlt&"' align=absmiddle>"
End Function

'//* 网站用户图例 *//
Public Function Img_User()
	Dim a,udim,ui
	'a="用户图例："
	udim=YSvoid.Cache_Get("user_group")
	For ui=0 To Ubound(udim,2)
		a=a&VbCrlf&"<img border=0 src='"&YSvoid.web_dir_skin&"Small/icon_"&udim(0,ui)&".gIf' align=absmiddle>&nbsp;"&udim(1,ui)
	Next
	a=a&VbCrlf&"<img border=0 src='"&YSvoid.web_dir_skin&"Small/icon_other.gIf' align=absmiddle>&nbsp;游客"
	If IsArray(udim) Then Erase udim
	Img_User=a
End Function

Function get_abc(var_s)
  '65-90 A-Z
  '97-122 a-z
  '48-58 0-9
  '45 -
  dim tmp,vars:vars=trim(var_s)
  If vars="" or isnull(vars) Then get_abc="-":exit Function
  vars=left(vars,1)
  tmp=int(asc(vars))
  If tmp>=48 and tmp<=57 Then get_abc=vars:exit Function
  If (tmp>=65 and tmp<=90) or (tmp>=97 and tmp<=122) Then get_abc=ucase(vars):exit Function
  tmp=tmp+65536
  If(tmp>=45217 and tmp<=45252) Then get_abc="A":exit Function
  If(tmp>=45253 and tmp<=45760) Then get_abc="B":exit Function
  If(tmp>=45761 and tmp<=46317) Then get_abc="C":exit Function
  If(tmp>=46318 and tmp<=46825) Then get_abc="D":exit Function
  If(tmp>=46826 and tmp<=47009) Then get_abc="E":exit Function
  If(tmp>=47010 and tmp<=47296) Then get_abc="F":exit Function
  If(tmp>=47297 and tmp<=47613) Then get_abc="G":exit Function
  If(tmp>=47614 and tmp<=48118) Then get_abc="H":exit Function
  If(tmp>=48119 and tmp<=49061) Then get_abc="J":exit Function
  If(tmp>=49062 and tmp<=49323) Then get_abc="K":exit Function
  If(tmp>=49324 and tmp<=49895) Then get_abc="L":exit Function
  If(tmp>=49896 and tmp<=50370) Then get_abc="M":exit Function
  If(tmp>=50371 and tmp<=50613) Then get_abc="N":exit Function
  If(tmp>=50614 and tmp<=50621) Then get_abc="O":exit Function
  If(tmp>=50622 and tmp<=50905) Then get_abc="P":exit Function
  If(tmp>=50906 and tmp<=51386) Then get_abc="Q":exit Function
  If(tmp>=51387 and tmp<=51445) Then get_abc="R":exit Function
  If(tmp>=51446 and tmp<=52217) Then get_abc="S":exit Function
  If(tmp>=52218 and tmp<=52697) Then get_abc="T":exit Function
  If(tmp>=52698 and tmp<=52979) Then get_abc="W":exit Function
  If(tmp>=52980 and tmp<=53640) Then get_abc="X":exit Function
  If(tmp>=53641 and tmp<=54480) Then get_abc="Y":exit Function
  If(tmp>=54481 and tmp<=62289) Then get_abc="Z":exit Function
  get_abc="-"
End Function

'//* 版权信息 *//
Public Sub Web_Copy()
	YSvoid.HtmlMain "copyright"
	YSvoid.HtmlRcod "closer",YSvoid.ModGet("closer")
	YSvoid.HtmlView(0)
End Sub

'================================================
'作  用：用户系统信息
'参  数：
'	svar : 用户系统信息
'================================================
Public Function User_Sys_Type(ByVal sVar)
	If YSvoid.Is_Null(sVar)="" Then
		User_Sys_Type="未知的系统信息"
		Exit Function
	End If
	Dim sys_dim,Temp1,sys2,s1,tmp
	sys_dim=Split(svar,";")
	If Int(Ubound(sys_dim))<2 Then
		User_Sys_Type="未知的系统信息"
		Exit Function
	End If
	sys2=sys_dim(2)
	s1=Len(sys2)
	tmp=Mid(sys2,s1,1)
	If tmp=")" Then sys2=Mid(sys2,1,s1-1)
	Temp1="操作系统："&sys2
	Temp1=Temp1&"，浏览器："&sys_dim(1)
	Temp1=Replace(Temp1,"MSIE","Internet Explorer")
	Temp1=Replace(Temp1,"NT 5.0","2000")
	Temp1=Replace(Temp1,"NT 5.1","XP")
	Temp1=Replace(Temp1,"NT 5.2","2003")
	If IsArray(sys_dim) Then Erase sys_dim
	User_Sys_Type=Temp1
End Function

Function Rank_Img(lTimes,tType)
 Dim Times,h_lTime,m_lTime,n_lTime,nh_lTime,nm_lTime,Rank_Temp,Rank,i,Temp,nRank_Time
 Rank_Temp=0
 h_lTime=lTimes \ 60
 m_lTime=lTimes Mod 60
 Do While h_lTime >= 5*Rank_Temp^2+15*Rank_Temp
  Rank_Temp=Rank_Temp+1
 Loop
 Rank=Rank_Temp-1
 n_lTime=(5*Rank_Temp^2+15*Rank_Temp)*60-lTimes
 nh_lTime=n_lTime \ 60
 nm_lTime=n_lTime Mod 60
 Times=Rank
 Do While Times > 0
  If Times \ 16 >0 Then
   For i=1 To Times \ 16
    Temp=Temp&"<img border=0 src="&YSvoid.Web_Dir&"images/Rank/sun.gIf>"
   Next
   Times=Times Mod 16
  End If
  If Times \ 4 >0 Then
   For i=1 to Times \ 4
    Temp=Temp&"<img border=0 src="&YSvoid.Web_Dir&"images/Rank/moon.gIf>"
   Next
   Times=Times Mod 4
  End If
  If Times \ 1 >0 Then
   For i=1 to Times \ 1
    Temp=Temp&"<img border=0 src="&YSvoid.Web_Dir&"images/Rank/star.gIf>"
   Next
   Times=Times Mod 1
  End If
 Loop
 Select Case Cint(tType)
 Case 1
   Rank_Img="<td colspan=2 class=gray><span title='Rank: "&Rank&"'>"&Temp&"</span>&nbsp;&nbsp;&nbsp;&nbsp;<font class=gray>今日在线<font class=red>"&today_onlinetime&"</font>分钟，总在线 <font class=red>"&lTimes&"</font> 分钟，<u><font class=tims>合计：<font class=red>"&h_lTime&"</font>小时<font class=red>"&m_lTime&"</font>分钟</font></u></td></tr>"
   Rank_Img=Rank_Img&"<tr class=bg_td><td></td><td colspan=2><font class=gray>离下次升级还有 <font class=red>"&n_lTime&"</font>分钟，<u><font class=tims>合计：<font class=red>"&nh_lTime&"</font>小时<font class=red>"&nm_lTime&"</font>分钟</font></font></u></td>"
 Case else
   Rank_Img= "<span title='Rank: "&Rank&"'>"&Temp&"</span>"
 End Select
End Function


'//* 格式化Sel_Type *//
Public Function Sel_Type(ByVal Seled,ByVal Selstr)
  Dim Snum,Sdim,Temp1
	If Selstr="" Then Exit Function
	If Instr(Selstr,"|")<0 then
	  Sel_Type="<option value='"&Selstr&"'>"&Selstr&"</option>"
	  Exit Function
	End If
	Sdim=Split(Selstr,"|")
	Snum=Ubound(Sdim)
	For i=0 To Snum
	  Temp1=Temp1&"<option value='"&Sdim(i)&"'"
	  If (Seled=Sdim(i)) Then Temp1=Temp1&" selected"
		Temp1=Temp1&">"&Sdim(i)&"</option>"
	Next
	Erase Sdim
  Sel_Type=Temp1
End Function

'//* 关键字显亮 *//
'--* strText：	 内容
'--* strKeyWord：关键字
'--* strText：	 关键字前部设置
'--* strKeyWord：关键字后部设置
Public Function KeyWord_Highlight(ByVal strText,ByVal strKeyWord,ByVal strBefore,ByVal strAfter)
    Dim nPos
    Dim nLen
    Dim nLenAll
    nLen = Len(strKeyWord)
    nLenAll = nLen + Len(strBefore) + Len(strAfter)
    KeyWord_Highlight = strText
    If nLen > 0 And Len(KeyWord_Highlight) > 0 Then
        nPos = InStr(1, KeyWord_Highlight, strKeyWord, 1)
        Do While nPos > 0
            KeyWord_Highlight = Left(KeyWord_Highlight, nPos - 1) & _
                strBefore & Mid(KeyWord_Highlight, nPos, nLen) & strAfter & _
                Mid(KeyWord_Highlight, nPos + nLen)
            nPos = InStr(nPos + nLenAll, KeyWord_Highlight, strKeyWord, 1)
        Loop
    End If
End Function

'//* 关键字链接 *//
'--* strText：	 内容
'--* strKeyWord：关键字
Public Sub KeyWord_Url(strText,strKeyWord)
    If Not YSvoid.KeyWordTrue or strKeyWord="" Then Exit Sub
	'-----------------挑选热门关键字------------------------
	Dim web_keyword,web_keywordhot,m,n
    web_keyword=split(strKeyWord,",")
    for m=0 to ubound(web_keyword)
	 for n=0 to ubound(YSvoid.ChannelKeyWord,2)
	   if YSvoid.ChannelKeyWord(4,n)=true and web_keyword(m)=YSvoid.ChannelKeyWord(1,n)then
	     web_keywordhot=web_keywordhot&web_keyword(m)&","
	   end if
	 next
	next
	'-------------------------------------------------------
    Dim stmp_keyword,stmpkeyword,stmp_url,j,Sii,web_keywords
	web_keywords=Replace(YSvoid.Web_Dim(4),"|",",")'站内关键字
	stmpkeyword=web_keywordhot&","&web_keywords
    stmp_keyword=split(stmpkeyword,",")
    For j=0 To ubound(stmp_keyword)
      For Sii=0 To ubound(YSvoid.ChannelKeyWord,2)
        If stmp_keyword(j)=YSvoid.ChannelKeyWord(1,Sii) Then
		  stmp_url=YSvoid.ChannelKeyWord(2,Sii)
		Else
		  stmp_url=""
		End If
		If stmp_url<>"" Then
		  If Left(stmp_url,7)<>"http://" Then
		    stmp_url="http://"&stmp_url
		  Else
		    stmp_url=stmp_url
		  End If
		Else
		  If Left(YSvoid.Web_Common(2),7)<>"http://" Then
		    stmp_url="http://"&YSvoid.Web_Common(2)
		  Else
		    stmp_url=YSvoid.Web_Common(2)
	      End If
		End If
	  Next
    strText=KeyWord_Highlight(strText,stmp_keyword(j),"<a href='"&stmp_url&"' target=_blank><font color=red>","</font></a>")
    Next
End Sub

Public Function Format_Remark_Var(ByVal strer,ByVal strt)
  If strer="" Or YSvoid.Is_Null(strer)="" Then Exit Function
  Dim re
  Set re=new RegExp
  re.IgnoreCase=True
  re.Global=True
  If strt>0 Then
    re.Pattern="\[align=right\](.|\n)*\[\/align\]"
    strer=re.Replace(strer,"")
  End If
  If strt>1 Then
    re.Pattern="\[quote\](.|\n)*\[\/quote\]"
    strer=re.Replace(strer,"")
  End If
  Set re=Nothing
  strer=YSvoid.Code_Health(strer)
  strer=YSvoid.Code_Word(strer)
  strer=Replace(strer,"'","\'")
  strer=Replace(strer,"<","&lt;")
  strer=Replace(strer,">","&gt;")
  strer=Replace(strer,vbcrlf,"\n")
  strer=Replace(strer,chr(10),"\n")
  strer=Replace(strer,chr(13),"")
  Format_Remark_Var=strer
End Function

%>