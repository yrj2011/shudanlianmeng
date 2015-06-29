<%

Class Cls_UbbCode
	Public Re,reed,isgetreed,Ubblists,MaxLoopcount,IsOpenHTML
	Private Sub Class_Initialize()
		MaxLoopcount=100  '循环的最多次数，避免死循环
		IsOpenHTML=0 '是否开放HTML支持
		'[/img]编号:1.[/upload]编号:2.[/dir]编号:3.[/qt]编号:4.[/mp]编号:5.
		'[/rm]编号:6.[/sound]编号:7.[/flash]编号:8.
		'[/url]编号:9.[/email]编号:10.
		'http,https,ftp,rtsp,mms编号:11.[/html]编号:12.
		'[/code]编号:13.[/color]编号:14.
		'[/face]编号:15.[/align]编号:16.
		'[/quote]编号:17.[/fly]编号:18.[/move]编号:19.
		'[/shadow]编号:20.[/glow]编号:21.[/size]编号:22.
		'[/i]编号:23.[/b]编号:24.[/u]编号:25.www编号:26.
		Ubblists=",1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,"
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
	End Sub
	Public Property Let OpenHTML(ByVal sValue)
		IsOpenHTML=Cint(sValue)
	End Property
	Private Sub class_terminate()
		Set Re=Nothing
	End Sub

	Public Function FilterJS(s)
		re.Pattern="(<s+cript(.[^>]*)>)"
		s=re.Replace(s,"&lt;&#83cript$2&gt;")
		re.Pattern="(<\/s+cript>)"
		s=re.Replace(s,"&lt;/&#83cript&gt;")
		re.Pattern = "(<IFRAME(.[^>]*)>)"
		s = re.Replace(s, "&lt;IFRAME$2&gt;")
		re.Pattern = "(<\/IFRAME>)"
		s = re.Replace(s, "&lt;/IFRAME&gt;")
		re.Pattern="(<body(.[^>]*)>)"
		s=re.Replace(s,"<body>")
		re.Pattern="(<\!(.[^>]*)>)"
		s=re.Replace(s,"&lt;$2&gt;")
		re.Pattern="(<\!)"
		s=re.Replace(s,"&lt;!")
		re.Pattern="(-->)"
		s=re.Replace(s,"--&gt;")
		re.Pattern="(javascript:)"
		s=re.Replace(s,"javascript:")
		re.Pattern="<((asp|\!|%))"
		s=re.Replace(s,"&lt;$1")
		re.Pattern="(>)("&vbNewLine&")(<)"
		s=re.Replace(s,"$1$3")
		re.Pattern="(>)("&vbNewLine&vbNewLine&")(<)"
		s=re.Replace(s,"$1$3")
		re.Pattern="(<br>)"
		s=re.Replace(s,"[br]")
		re.Pattern="<(\w+)(&nbsp;)+([^>]*)>"
		s=re.Replace(s,"<$1 $3>")
		re.Pattern="\[(br)\]"
		s=re.Replace(s,"<$1>")
		re.Pattern="(jscript:)"
		s=re.Replace(s,"&#106script:")
		re.Pattern="(js:)"
		s=re.Replace(s,"&#106s:")
		re.Pattern="(about:)"
		s=re.Replace(s,"about&#58")
		re.Pattern="(file:)"
		s=re.Replace(s,"file&#58")
		re.Pattern="(document.cookie)"
		s=re.Replace(s,"documents&#46cookie")
		re.Pattern="(vbscript:)"
		s=re.Replace(s,"&#118bscript:")
		re.Pattern="(vbs:)"
		s=re.Replace(s,"&#118bs:")
		re.Pattern="(on(mouse|exit|error|click|key))"
		s=re.Replace(s,"&#111n$2")
		FilterJS=s
	End Function

	Public Function FilterEm(s)
		if YSvoid.Is_Null(s)="" then
		  FilterEm=""
		  exit function
		End If
		dim emci
		if instr(s,"[emb")>0 then
		  for emci=1 to 13
		    s=Replace(s,"[emb"&emci&"]","<image src="&YSvoid.Web_Dir&"Images/em/emb/"&emci&".gif border=0>&nbsp;")
		  next
		End If
		if instr(s,"[emc")>0 then
		  for emci=1 to 25
		    s=Replace(s,"[emc"&emci&"]","<image src="&YSvoid.Web_Dir&"Images/em/emc/"&emci&".gif border=0>&nbsp;")
		  next
		End If
		if instr(s,"[emq")>0 then
		  for emci=1 to 80
		    s=Replace(s,"[emq"&emci&"]","<image src="&YSvoid.Web_Dir&"Images/em/emq/"&emci&".gif border=0>&nbsp;")
		  next
		End If
		if instr(s,"[emm")>0 then
		  for emci=1 to 67
		    s=Replace(s,"[emm"&emci&"]","<image src="&YSvoid.Web_Dir&"Images/em/emm/"&emci&".gif border=0>&nbsp;")
		  next
		End If
		if instr(s,"[em")>0 then
		  for emci=1 to 16
		    s=Replace(s,"[em"&emci&"]","<image src="&YSvoid.Web_Dir&"Images/em/em/"&emci&".gif border=0>&nbsp;")
		  next
		End If
		FilterEm=s
	End Function

	Public Function UbbCode(s)
		Dim ii,ranNum
		Dim po
		s=YSvoid.Code_Health(s)
		s=FilterJS(s)
	    'Html标签替换
		If IsOpenHTML=0 Then
        s=Replace(s,"<","&lt;")
        s=Replace(s,">","&gt;")
        End If
		s=FilterEm(s)
		s=Replace(s, vbNewLine, "<br>")
		s=Replace(s, CHR(10), "")
		s=Replace(s, CHR(13), "")
		s=Replace(s,">"&CHR(10)&"<","><")
		s=Replace(s,">"&CHR(10)&CHR(10)&"<","><")
		s=Replace(s,CHR(10),"<br>")
		s=Replace(s, "  ", "&nbsp;")
		s=Replace(s, "  ", "&nbsp;&nbsp;")
		s=Replace(s,"[list]","<ul>")
		s=Replace(s,"[list=1]","<ol type=1>")
		s=Replace(s,"[list=a]","<ol type=a>")
		s=Replace(s,"[list=A]","<ol type=A>")
		s=Replace(s,"[*]","<li>")
		s=Replace(s,"[/list]","</ul></ol>")
		'去掉图片中的脚本代码
		re.Pattern="<IMG.[^>]*SRC(=| )(.[^>]*)>"
		s=re.replace(s,"<IMG SRC=$2 onclick=""javascript:window.open(this.src);"" style=""CURSOR: pointer"">")
		'img code
		If InStr(Ubblists,",1,")>0 Then s=YSvoid_UbbCode_S1(s,"\[IMG\]","\[\/IMG\]","IMG","<a onfocus=this.blur() href=""$1"" target=_blank title=新窗口打开><IMG SRC=""$1"" border=0 ></a>")
		'upload code
		If InStr(Ubblists,",2,")>0 Then s=YSvoid_UbbCode_U(s,1) '是否开放标签

		'media code
		If InStr(Ubblists,",3,")>0 Then s=YSvoid_UbbCode_iS2(s,"\[DIR=(.[^\[]*)\]","\[\/DIR\]","DIR","<object classid=clsid:166B1BCA-3F9C-11CF-8075-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=7,0,2,0 width=$1 height=$2><param name=src value=$3><embed src=$3 pluginspage=http://www.macromedia.com/shockwave/download/ width=$1 height=$2></embed></object>","<a href=$3 target=_blank>$3</a>",1,"=*([0-9]*),*([0-9]*)")

		If InStr(Ubblists,",4,")>0 Then s=YSvoid_UbbCode_iS2(s,"\[QT=(.[^\[]*)\]","\[\/QT\]","QT","<embed src=$3 width=$1 height=$2 autoplay=true loop=false controller=true playeveryframe=false cache=false scale=TOFIT bgcolor=#000000 kioskmode=false targetcache=false pluginspage=http://www.apple.com/quicktime/>","<a href=$3 target=_blank>$3</a>",1,"=*([0-9]*),*([0-9]*)")

		If InStr(Ubblists,",5,")>0 Then
		s=YSvoid_UbbCode_iS2(s,"\[MP=(.[^\[]*)\]","\[\/MP\]","MP","<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=$1 height=$2 ><param name=ShowStatusBar value=-1><param name=Filename value=$3><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=""$3"" width=$1 height=$2></embed></object>","<a href=""$3"" target=_blank>$3</a>",1,"=*([0-9]*),*([0-9]*)")
		'MediaPlayer自定义播放模式；
		s=YSvoid_UbbCode_iS2(s,"\[MP=(.[^\[]*)\]","\[\/MP\]","MP","<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=$1 height=$2 ><PARAM NAME=AUTOSTART VALUE=$3 ><param name=ShowStatusBar value=-1><param name=Filename value=$4><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=""$4"" width=$1 height=$2></embed></object>","<a href=""$4"" target=_blank>$4</a>",1,"=*([0-9]*),*([0-9]*),*([0|1|true|false]*)")
		End If

		If InStr(Ubblists,",6,")>0 Then
		s=YSvoid_UbbCode_iS2(s,"\[RM=(.[^\[]*)\]","\[\/RM\]","RM","<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=$1 height=$2><PARAM NAME=SRC VALUE=""$3""><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=$1><PARAM NAME=SRC VALUE=""$3""><PARAM NAME=AUTOSTART VALUE=-1><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1></OBJECT>","<a href=$3 target=_blank>$3</a>",1,"=*([0-9]*),*([0-9]*)")
		'RealPlayer自定义播放模式；
		s=YSvoid_UbbCode_iS2(s,"\[RM=(.[^\[]*)\]","\[\/RM\]","RM","<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=$1 height=$2><PARAM NAME=SRC VALUE=""$4""><PARAM NAME=CONSOLE VALUE=""$4""><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=$3 ></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=""video"" width=$1><PARAM NAME=SRC VALUE=""$4""><PARAM NAME=AUTOSTART VALUE=$3><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=""$4""></OBJECT>","<a href=$4 target=_blank>$4</a>",1,"=*([0-9]*),*([0-9]*),*([0|1|true|false]*)")
		End If

		If InStr(Ubblists,",7,")>0 Then s=YSvoid_UbbCode_S2(s,"\[sound\]","\[\/sound\]","sound","<a href=""$1"" target=_blank><IMG SRC="&YSvoid.Web_Dir&"Forum/Images/FileType/mid.gif border=0 alt=""背景音乐""></a><bgsound src=""$1"" loop=""-1"">","<a href=$1 target=_blank>$1</a>",1)

		'flash code
		If InStr(Ubblists,",8,")>0 Then
			s=YSvoid_UbbCode_S2(s,"\[FLASH\]","\[\/FLASH\]","FLASH","<a href=""$1"" TARGET=_blank><IMG SRC="&YSvoid.Web_Dir&"Forum/Images/FileType/swf.gif border=0 alt=""点击开新窗口欣赏该FLASH动画!"">[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE=""$1""><PARAM NAME=quality VALUE=high><embed src=""$1"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$1</embed></OBJECT>","<IMG SRC="&YSvoid.Web_Dir&"Forum/Images/FileType/swf.gif border=0><a href=$1 target=_blank>$1</a>",1)
			s=YSvoid_UbbCode_iS2(s,"\[FLASH=(.[^\[]*)\]","\[\/FLASH\]","FLASH","<a href=""$3"" TARGET=_blank><IMG SRC="&YSvoid.Web_Dir&"Forum/Images/FileType/swf.gif border=0 alt=""点击开新窗口欣赏该FLASH动画!"">[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$1 height=$2><PARAM NAME=movie VALUE=""$3""><PARAM NAME=quality VALUE=high><embed src=""$3"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$1 height=$2>$3</embed></OBJECT>","<a href=$3 target=_blank>$3</a>",1,"=*([0-9]*),*([0-9]*)")
		End If
		'flv code
		If InStr(Ubblists,",27,")>0 Then
		    s=YSvoid_UbbCode_S2(s,"\[FLV\]","\[\/FLV\]","FLV","<OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE="&YSvoid.Web_Dir&"images/small/video_flv.swf><PARAM NAME=quality VALUE=high><PARAM name=""allowFullScreen"" value=""true""><PARAM name=""FlashVars"" value=""vcastr_file='+""$1""+'&vcastr_title='+texts+'""><embed Src="&YSvoid.Web_Dir&"images/small/video_flv.swf allowFullScreen=""true"" FlashVars=""vcastr_file='+""$1""+'&vcastr_title='+texts+'"" menu=""false"" quality=""high"" width=500 height=400 type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer""></OBJECT>","",1)
			s=YSvoid_UbbCode_iS2(s,"\[FLV=(.[^\[]*)\]","\[\/FLV\]","FLV","<OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$1 height=$2><PARAM NAME=movie VALUE="&YSvoid.Web_Dir&"images/small/video_flv.swf><PARAM NAME=quality VALUE=high><PARAM name=""allowFullScreen"" value=""true""><PARAM name=""FlashVars"" value=""vcastr_file=$3&isAutoPlay=1&IsContinue=1&LogoUrl="&YSvoid.Web_Dir&"images/Flv_logo.png&BarColor=0x1D7D8B&BarPosition=1""><embed Src="&YSvoid.Web_Dir&"images/small/video_flv.swf allowFullScreen=""true"" FlashVars=""vcastr_file=$3&isAutoPlay=1&IsContinue=1&LogoUrl="&YSvoid.Web_Dir&"images/Flv_logo.png&BarColor=0x1D7D8B&BarPosition=1"" menu=""false"" quality=""high"" width=$1 height=$2 type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer""></OBJECT>","<a href=$3 target=_blank>$3</a>",1,"=*([0-9]*),*([0-9]*)")
		End If
		'=========================================================
		'url code
		If InStr(Ubblists,",9,")>0 Then
			s=YSvoid_UbbCode_S1(s,"\[URL\]","\[\/URL\]","URL","<img align=absmiddle src="&YSvoid.web_dir_skin&"small/url.gif><a href=""$1"" target=_blank>$1</a>")
			're.Pattern="(\[URL=(.[^:\/\/|\[]*)\])(.[^\[]*)(\[\/URL\])"
			's= re.Replace(s,"<A HREF=""http://$2"" TARGET=_blank>$3</A>")
			re.Pattern="(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
			s= re.Replace(s,"<a href=""$2"" target=_blank>$3</a>")
			's= re.Replace(s,"<img align=absmiddle src="&YSvoid.web_dir_skin&"small/url.gif><A HREF=""$2"" TARGET=_blank>$3</A>")
		End If
		'email code
		If InStr(Ubblists,",10,")>0 Then
			s=YSvoid_UbbCode_S1(s,"\[EMAIL\]","\[\/EMAIL\]","EMAIL","<img align=absmiddle src="&YSvoid.web_dir_skin&"small/email.gif ><A HREF=""mailto:$1"">$1</A>")
			re.Pattern="(\[EMAIL=(\S+\@.[^\[]*)\])(.[^\[]*)(\[\/EMAIL\])"
			s= re.Replace(s,"<img align=absmiddle src="&YSvoid.web_dir_skin&"small/email.gif ><A HREF=""mailto:$2"" TARGET=_blank>$3</A>")
		End If

		If InStr(Ubblists,",12,")>0 Then s=YSvoid_UbbCode_S1(s,"\[HTML\]","\[\/HTML\]","HTML","<table border=0 width=""100%"" align=center><tr><td>以下内容为程序代码:$1</td></tr></table>")
		If InStr(Ubblists,",13,")>0 Then s=YSvoid_UbbCode_S1(s,"\[code\]","\[\/code\]","code","<table border=0 width=""100%"" align=center><tr><td id=ubb_code>$1</td></tr></table>")
		If InStr(Ubblists,",14,")>0 Then s=YSvoid_UbbCode_C(s)
		If InStr(Ubblists,",15,")>0 Then s=YSvoid_UbbCode_F(s)
		If InStr(Ubblists,",16,")>0 Then s=YSvoid_UbbCode_Align(s)
		If InStr(Lcase(s),"center]")>0 Then s=YSvoid_UbbCode_S1(s,"\[center\]","\[\/center\]","center","<center>$1</center>")
		If InStr(Ubblists,",17,")>0 Then s=YSvoid_UbbCode_Q(s)
		If InStr(Ubblists,",18,")>0 Then s=YSvoid_UbbCode_S1(s,"\[fly\]","\[\/fly\]","fly","<marquee width=100% behavior=alternate scrollamount=3>$1</marquee>")
		If InStr(Ubblists,",19,")>0 Then s=YSvoid_UbbCode_S1(s,"\[move\]","\[\/move\]","move","<MARQUEE scrollamount=3>$1</marquee>")
		If InStr(Ubblists,",20,")>0 Then s=YSvoid_UbbCode_iS1(s,"\[SHADOW=(.[^\[]*)\]","\[\/SHADOW\]","SHADOW","<div style=""width:$1px;filter:shadow(color=$2, strength=$3)"">$4</div>","=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)")
		If InStr(Ubblists,",21,")>0 Then s=YSvoid_UbbCode_iS1(s,"\[GLOW=(.[^\[]*)\]","\[\/GLOW\]","GLOW","<div style=""width:$1px;filter:glow(color=$2, strength=$3)"">$4</div>","=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)")
		If InStr(Ubblists,",22,")>0 Then s=YSvoid_UbbCode_S(s)
		If InStr(Ubblists,",23,")>0 Then s=YSvoid_UbbCode_S1(s,"\[i\]","\[\/i\]","i","<i>$1</i>")
		If InStr(Ubblists,",24,")>0 Then s=YSvoid_UbbCode_S1(s,"\[b\]","\[\/b\]","b","<b>$1</b>")
		If InStr(Ubblists,",25,")>0 Then s=YSvoid_UbbCode_S1(s,"\[u\]","\[\/u\]","u","<u>$1</u>")
		'不开放HTML支持，不转换HREF
		If IsOpenHTML=1 Then
			'自动识别网址
			If InStr(Ubblists,",11,")>0 Then
				re.Pattern = "^((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]:+!]+([^<>""|'])+)"
				s = re.Replace(s,"<a target=_blank href=$1>$1</a>")
				re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]:+!]+([^<>""])+)$([^\[|']*)"
				s = re.Replace(s,"<a target=_blank href=$1>$1</a>")
				re.Pattern = "([^>=""])((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]:+!]+([^<>""|'])+)"
				s = re.Replace(s,"$1<a target=_blank href=$2>$2</a>")
			End If
			'自动识别www等开头的网址
			If InStr(Ubblists,",26,")>0 Then
				re.Pattern = "([\s])((www|cn)[.](\w)+[.]{1,}(net|com|cn|org|cc)(((\/[\~]*|\\[\~]*)(\w)+)|[.](\w)+)*(((([?](\w)+){1}[=]*))*((\w)+){1}([\&](\w)+[\=](\w)+)*)*)"
				s = re.Replace(s,"<a target=_blank href=""http://$2"">$2</a>")
			End If
		End If
		UbbCode=bbimg(s,YSvoid.Format_Mid_Num(24))
	End Function

	Private Function bbimg(strText,ssize)
		Dim s
		s=strText
		re.Pattern="<img(.[^>]*)>"
		'If ssize=500 Then
			's=re.replace(s,"<img$1 onmousewheel=""return bbimg(this)"" onload=""javascript:if(this.width>screen.width-"&ssize&")this.style.width=screen.width-"&ssize&";"">")
			s=re.replace(s,"<img$1 onmousewheel=""return bbimg(this)"" onload=""javascript:if(this.width>screen.width-"&ssize&"){this.style.width=screen.width-"&ssize&"};"">")
		'Else
			's=re.replace(s,"<img$1 onmousewheel=""return bbimg(this)"" onload=""javascript:if(this.width>screen.width-"&ssize&")this.style.width=screen.width-"&ssize&";if(this.height>250)this.style.width=(this.width*250)/this.height;"">")
		'End If
		bbimg=s
	End Function

	Private Function YSvoid_UbbCode_S1(strText,uCodeL,uCodeR,uCodeC,tCode)
		Dim s
		s=strText
		re.Pattern=uCodeL
		s=re.replace(s, chr(1) & uCodeC & chr(2))
		re.Pattern=uCodeR
		s=re.replace(s, chr(1) & "/" & uCodeC & chr(2))
		re.Pattern="\x01"&uCodeC&"\x02\x01\/"&uCodeC&"\x02"
		s=re.Replace(s,"")
		re.Pattern="\x01"&uCodeC&"\x02(.[^\x01]*)\x01\/"&uCodeC&"\x02"
		s=re.Replace(s,tCode)
		re.Pattern="\x02"
		s=re.replace(s, "]")
		re.Pattern="\x01"
		s=re.replace(s, "[")
		YSvoid_UbbCode_S1=s
	End Function

	Private Function YSvoid_UbbCode_iS1(strText,uCodeL,uCodeR,uCodeC,tCode,iCode)
		Dim s
		s=strText
		re.Pattern=uCodeL
		s=re.replace(s, chr(1) & uCodeC & "=$1" & chr(2))
		re.Pattern=uCodeR
		s=re.replace(s, chr(1) & "/" & uCodeC & chr(2))
		re.Pattern="\x01"&uCodeC&iCode&"\x02\x01\/"&uCodeC&"\x02"
		s=re.Replace(s,"")
		re.Pattern="\x01"&uCodeC&iCode&"\x02(.[^\x01]*)\x01\/"&uCodeC&"\x02"
		s=re.Replace(s,tCode)
		re.Pattern="\x02"
		s=re.replace(s, "]")
		re.Pattern="\x01"
		s=re.replace(s, "[")
		YSvoid_UbbCode_iS1=s
	End Function


	Private Function YSvoid_UbbCode_S2(strText,uCodeL,uCodeR,uCodeC,tCode1,tCode2,Flag)
		Dim s
		s=strText
		re.Pattern=uCodeL
		s=re.replace(s, chr(1) & uCodeC & chr(2))
		re.Pattern=uCodeR
		s=re.replace(s, chr(1) & "/" & uCodeC & chr(2))
		re.Pattern="\x01"&uCodeC&"\x02(.[^\x01]*)\x01\/"&uCodeC&"\x02"
		If Flag = 1 Then
			s=re.Replace(s,tCode1)
		Else
			s=re.Replace(s,tCode2)
		End If
		re.Pattern="\x02"
		s=re.replace(s, "]")
		re.Pattern="\x01"
		s=re.replace(s, "[")
		YSvoid_UbbCode_S2=s
	End Function

	Private Function YSvoid_UbbCode_iS2(strText,uCodeL,uCodeR,uCodeC,tCode1,tCode2,Flag,iCode)
		Dim s
		s=strText
		re.Pattern=uCodeL
		s=re.replace(s, chr(1) & uCodeC & "=$1" & chr(2))
		re.Pattern=uCodeR
		s=re.replace(s, chr(1) & "/" & uCodeC & chr(2))
		Rem 过滤媒体标签中的FLASH
		If Instr(uCodeC,"FLASH") = 0 Then
			re.Pattern="(.(swi))"
			s=re.Replace(s,"")
		End If
		re.Pattern="\x01"&uCodeC&iCode&"\x02(.[^\x01]*)\x01\/"&uCodeC&"\x02"
		If Flag = 1 Then
			s=re.Replace(s,tCode1)
		Else
			s=re.Replace(s,tCode2)
		End If
		re.Pattern="\x02"
		s=re.replace(s, "]")
		re.Pattern="\x01"
		s=re.replace(s, "[")
		YSvoid_UbbCode_iS2=s
	End Function

	Private Function YSvoid_UbbCode_Align(strText)
		Dim s
		s=strText
		re.Pattern="\[ALIGN=(center|left|right)\]"
		s=re.replace(s, chr(1) & "ALIGN=$1" & chr(2))
		re.Pattern="\[\/ALIGN\]"
		s=re.replace(s, chr(1) & "/ALIGN" & chr(2))
		re.Pattern="\x01ALIGN=(center|left|right)\x02\x01\/ALIGN\x02"
		s=re.Replace(s,"")
		re.Pattern="\x01ALIGN=(center|left|right)\x02(.[^\x01]*)\x01\/ALIGN\x02"
		s=re.Replace(s,"<div align=$1>$2</div>")
		re.Pattern="\x02"
		s=re.replace(s, "]")
		re.Pattern="\x01"
		s=re.replace(s, "[")
		YSvoid_UbbCode_Align=s
	End Function

	Private Function YSvoid_UbbCode_U(strText,Flag)	'(内容，是否开放图片标签)
		Dim s
		s=strText
		re.Pattern="\[UPLOAD=(gif|jpg|jpeg|bmp|png|swf|swi)\]"
		s=re.replace(s, chr(1) & "UPLOAD=$1" & chr(2))
		re.Pattern="\[\/UPLOAD\]"
		s=re.replace(s, chr(1) & "/UPLOAD" & chr(2))
		re.Pattern="\x01UPLOAD=(gif|jpg|jpeg|bmp|png)\x02\x01\/UPLOAD\x02"
		s=re.Replace(s,"")
		re.Pattern="\x01UPLOAD=(gif|jpg|jpeg|bmp|png)\x02(.[^\x01]*)\x01\/UPLOAD\x02"
		If Flag = 1 Then
			s= re.Replace(s,"<br><IMG SRC="""&YSvoid.Web_Dir&"Forum/Images/FileType/$1.gif"" border=0 >相关图片如下：<br><A HREF=""$2"" TARGET=_blank id=""ImgSpan""><IMG SRC=""$2"" border=0 alt=按此在新窗口浏览图片 ></A>")
		Else
			s= re.Replace(s,"<br><IMG SRC="""&YSvoid.Web_Dir&"Forum/Images/FileType/$1.gif"" border=0 ><A HREF=""$2"" TARGET=_blank>$2</A>")
		End If
		re.Pattern="\x01UPLOAD=(swf|swi)\x02(.[^\x01]*)\x01\/UPLOAD\x02"
		If Flag = 1 Then
			s= re.Replace(s,"<br><IMG SRC="""&YSvoid.Web_Dir&"Forum/Images/FileType/swf.gif"" border=0 ><A HREF=""$2"" TARGET=_blank>点击浏览该FLASH文件</A>：<br><embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=300></embed>")
		Else
			s= re.Replace(s,"<br><IMG SRC="""&YSvoid.Web_Dir&"Forum/Images/FileType/swf.gif"" border=0 ><A HREF=""$2"" TARGET=_blank>$2</A>")
		End If
		re.Pattern="\x02"
		s=re.replace(s, "]")
		re.Pattern="\x01"
		s=re.replace(s, "[")
		re.Pattern="\[UPLOAD=(.[^\[]*)\]"
		s=re.replace(s, chr(1) & "UPLOAD=$1" & chr(2))
		re.Pattern="\[\/UPLOAD\]"
		s=re.replace(s, chr(1) & "/UPLOAD" & chr(2))
		re.Pattern="\x01UPLOAD=(.[^\x01]*)\x02\x01\/UPLOAD\x02"
		s=re.Replace(s,"")
		re.Pattern="\x01UPLOAD=(.[^\x01]*)\x02(.[^\x01]*)\x01\/UPLOAD\x02"
		s= re.Replace(s,"<br><IMG SRC="""&YSvoid.Web_Dir&"Forum/Images/FileType/$1.gif"" border=0> <a href=""$2"" target=_blank>点击下载</a><br>")
		's= re.Replace(s,"<br><IMG SRC="""&YSvoid.Web_Dir&"Forum/Images/FileType/$1.gif"" border=0> <a href=""$2"" target=_blank>点击浏览该文件</a><br><IMG src=""$2"" border=0 >")
		re.Pattern="\x02"
		s=re.replace(s, "]")
		re.Pattern="\x01"
		s=re.replace(s, "[")
		YSvoid_UbbCode_U=s
	End Function

	Private Function YSvoid_UbbCode_S(strText)
		Dim s
		Dim Test
		Dim LoopCount
		LoopCount=0
		s=strText
		Do While True
			re.Pattern="\[SIZE=([1-7])\]"
			Test=re.Test(s)
			If Test Then
				re.Pattern="\[\/SIZE\]"
				Test=re.Test(s)
				If Test Then
					re.Pattern="\[SIZE=([1-7])\]"
					s=re.replace(s, chr(1) & "SIZE=$1" & chr(2))
					re.Pattern="\[\/SIZE\]"
					s=re.replace(s, chr(1) & "/SIZE" & chr(2))
					re.Pattern="\x01SIZE=([1-7])\x02\x01\/SIZE\x02"
					s=re.Replace(s,"")
					re.Pattern="\x01SIZE=([1-7])\x02(.[^\x01]*)\x01\/SIZE\x02"
					s=re.Replace(s,"<font size=$1>$2</font>")
					re.Pattern="\x02"
					s=re.replace(s, "]")
					re.Pattern="\x01"
					s=re.replace(s, "[")
				Else
					Exit Do
				End If
			Else
				Exit Do
			End If
			 LoopCount=LoopCount+1
			 If LoopCount>MaxLoopCount Then Exit Do
		Loop
		YSvoid_UbbCode_S=s
	End Function

	Private Function YSvoid_UbbCode_Q(strText)
		Dim s
		Dim Test
		Dim LoopCount
		LoopCount=0
		s=strText
		re.Pattern="\[QUOTE\]"
		Test=re.Test(s)
		If Test Then
			re.Pattern="\[\/QUOTE\]"
			Test=re.Test(s)
			If Test Then
				re.Pattern="\[QUOTE\]"
				s=re.replace(s, chr(1) & "QUOTE" & chr(2))
				re.Pattern="\[\/QUOTE\]"
				s=re.replace(s, chr(1) & "/QUOTE" & chr(2))
				Do
					re.Pattern="\x01QUOTE\x02\x01\/QUOTE\x02"
					s=re.Replace(s,"")
					re.Pattern="\x01QUOTE\x02(.[^\x01]*)\x01\/QUOTE\x02"
					s=re.Replace(s,"<table border=0 width=""100%"" align=center><tr><td id=ubb_quote>$1</td></tr></table>")
					Test=re.Test(s)
					LoopCount=LoopCount+1
					If LoopCount>MaxLoopCount Then Exit Do
				Loop While(Test)
				re.Pattern="\x02"
				s=re.replace(s, "]")
				re.Pattern="\x01"
				s=re.replace(s, "[")
			End If
		End If
		YSvoid_UbbCode_Q=s
	End Function

	Private Function YSvoid_UbbCode_C(strText)
		Dim s
		Dim Test
		Dim LoopCount
		LoopCount=0
		s=strText
		Do While True
			re.Pattern="\[COLOR=(.[^\[]*)\]"
			Test=re.Test(s)
			If Test Then
				re.Pattern="\[\/COLOR\]"
				Test=re.Test(s)
				If Test Then
					re.Pattern="\[COLOR=(.[^\[]*)\]"
					s=re.replace(s, chr(1) & "COLOR=$1" & chr(2))
					re.Pattern="\[\/COLOR\]"
					s=re.replace(s, chr(1) & "/COLOR" & chr(2))
					re.Pattern="\x01COLOR=(.[^\x01]*)\x02\x01\/COLOR\x02"
					s=re.Replace(s,"")
					re.Pattern="\x01COLOR=(.[^\x01]*)\x02(.[^\x01]*)\x01\/COLOR\x02"
					s=re.Replace(s,"<font color=$1>$2</font>")
					re.Pattern="\x02"
					s=re.replace(s, "]")
					re.Pattern="\x01"
					s=re.replace(s, "[")
				Else
					Exit Do
				End If
			Else
				Exit Do
			End If
			LoopCount=LoopCount+1
			If LoopCount>MaxLoopCount Then Exit Do
		Loop
		YSvoid_UbbCode_C=s
	End Function

	Private Function YSvoid_UbbCode_F(strText)
		Dim s
		Dim Test
		Dim LoopCount
		LoopCount=0
		s=strText
		Do While True
			re.Pattern="\[FACE=(.[^\[]*)\]"
			Test=re.Test(s)
			If Test Then
				re.Pattern="\[\/FACE\]"
				Test=re.Test(s)
				If Test Then
					re.Pattern="\[FACE=(.[^\[]*)\]"
					s=re.replace(s, chr(1) & "FACE=$1" & chr(2))
					re.Pattern="\[\/FACE\]"
					s=re.replace(s, chr(1) & "/FACE" & chr(2))
					re.Pattern="\x01FACE=(.[^\x01]*)\x02\x01\/FACE\x02"
					s=re.Replace(s,"")
					re.Pattern="\x01FACE=(.[^\x01]*)\x02(.[^\x01]*)\x01\/FACE\x02"
					s=re.Replace(s,"<font face=$1>$2</font>")
					re.Pattern="\x02"
					s=re.replace(s, "]")
					re.Pattern="\x01"
					s=re.replace(s, "[")
				Else
					Exit Do
				End If
			Else
				Exit Do
			End If
			LoopCount=LoopCount+1
			If LoopCount>MaxLoopCount Then Exit Do
		Loop
		YSvoid_UbbCode_F=s
	End Function
End Class

%>