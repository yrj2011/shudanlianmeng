<%
'****************************
'ϵͳԤ������
'****************************
Class System_Cls
	Private LocalCacheName,Cache_Data
	Public Reloadtime,CacheName,CacheData,savelog,SqlQueryNum '��������
	Public pNum,pNum2

	'����System_Cls��Ԥ��������
	Private Sub Class_Initialize()
		Dim UserAgent,web_CacheName
		web_CacheName = "Jxqgw"   '�������ƣ����һ��վ���ж��վ����ĳɲ�ͬ����
		UserAgent = Trim(Lcase(Request.Servervariables("HTTP_USER_AGENT")))
		If InStr(UserAgent,"teleport") > 0 or InStr(UserAgent,"webzip") > 0 or InStr(UserAgent,"flashget")>0 or InStr(UserAgent,"offline")>0 Then
			Response.Write "�벻Ҫ����teleport/Webzip/Flashget/Offline�ȹ����������վ��"
			Response.End
		End If
		CacheName=Replace(Server.MapPath("\index.asp"),"index.asp","")
		'Response.Write(CacheName)
		'ssssResponse.End()
		if right(CacheName,3)="inc" then
			CacheName=Replace(CacheName,"inc","")
		end if
		CacheName=Replace(CacheName,":","")
		CacheName=Replace(CacheName,"\","")	'�ش���󣬷�������	
		CacheName=CacheName & web_CacheName  '���
		Reloadtime=14400
		savelog=0
		SqlQueryNum=0
		pNum=1:pNum2=0
	End Sub

	'����System_Cls����ֹ��������
	Private Sub class_terminate()
		If IsObject(Conn) Then 
			Conn.Close
			Set Conn = Nothing
		End If 
	End Sub

	'Cache�������
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



	'����ϵͳ��Դ����
	Public Site_Info,Site_Setting,Site_Version,Site_Copyright,BadWords,rBadWord
	'ȡ��ϵͳ������Դ
	Public Sub GetSite_Setting()
		Name="setup"
		If ObjIsEmpty() Then ReloadSetup()
		CacheData=value

		'ÿ�ո�������
		Name="Date"
		'��һ��������վ��������IIS��ʱ����ػ���
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


  '��������ر���
	  
		 

	 
  '�Ѿ�ɾ��

	'ҳ����ʾ�ຯ��
	  '�Ѿ�ɾ��


'ҳ����ʾ���ݺ���
	  '�Ѿ�ɾ��

	'��ȡ��վ���ʼ���
    'DELETE

	Public sub loadVote()
		dim Rs,SQL,tmpdata,i
		Name="vote"&ChannelID
		SQL="select top 1 * from Vote where IsSelected=True and (ChannelID=0 or ChannelID=" & ChannelID & ") order by ID Desc"
		Set Rs= execute(SQL)
		if Rs.bof and Rs.eof then 
			tmpdata = "<font color='#ff9900'>��&nbsp;</font>û���κε���" 
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
'����������Ŀ '��֪����ǰ̨�õĻ���
'=================================================
'��������ArticleContentshiyu
'��  �ã���ʾ�������ԡ����⡢���ߡ��������ڡ����������Ϣ
'��  ����intTitleLen  ----��������ַ�����һ������=����Ӣ���ַ�
'        ShowProperty ----�Ƿ���ʾ�������ԣ��̶�/�Ƽ�/��ͨ����TrueΪ��ʾ��FalseΪ����ʾ
'        ShowIncludePic ---�Ƿ���ʾ��[ͼ��]��������TrueΪ��ʾ��FalseΪ����ʾ
'        ShowAuthor -------�Ƿ���ʾ�������ߣ�TrueΪ��ʾ��FalseΪ����ʾ
'        ShowDateType -----��ʾ�������ڵ���ʽ��0Ϊ����ʾ��1Ϊ��ʾ�����գ�2Ϊֻ��ʾ���ա�
'        ShowHits ---------�Ƿ���ʾ���µ������TrueΪ��ʾ��FalseΪ����ʾ
'        ShowHot ----------�Ƿ���ʾ�������±�־��TrueΪ��ʾ��FalseΪ����ʾ
'=================================================
function ArticleContentshiyu(intTitleLen,ShowProperty,ShowIncludePic,ShowAuthor,ShowDateType,ShowHits,ShowHot)
   	dim i,strTemp,TitleStr,Author,AuthorName,AuthorEmail
    	i=0
	strTemp="<table width=100%  border=0 cellpadding=0 cellspacing=0>"
	do while not rsArticle.eof
		strTemp=strTemp & "<tr><td>"
		if ShowProperty=True then
			if rsArticle("OnTop")=true then
				strTemp = strTemp & "<img src='images/article_ontop.gif' alt='�̶�����'>&nbsp;"
			elseif rsArticle("Elite")=true then
				strTemp = strTemp & "<img src='images/article_elite.gif' alt='�Ƽ�����'>&nbsp;"
			else
				strTemp = strTemp & "<img src='images/article_common.gif' alt='��ͨ����'>&nbsp;"
			end if
		end if
		if ShowIncludePic=True and rsArticle("IncludePic")=true then
			strTemp = strTemp & "<font color=blue>[ͼ��]</font>"
		end if
		Author=rsArticle("Author")
		if instr(Author,"|")>0 then
			AuthorName=left(Author,instr(Author,"|")-1)
			AuthorEmail=right(Author,len(Author)-instr(Author,"|")-1)
		else
			AuthorName=Author
			AuthorEmail=""
		end if
		strTemp = strTemp & "<a href='" & rsArticle("LayoutFileName") & "?ArticleID=" & rsArticle("articleid") & "' title='���±��⣺" & rsArticle("Title") & vbcrlf & "�������ߣ�" & AuthorName & vbcrlf & "����ʱ�䣺" & rsArticle("UpdateTime") & vbcrlf & "���������" & rsArticle("Hits") & "' target='_blank'>"
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
			strTemp= strTemp & "<img src='images/hot.gif' alt='�ȵ�����'>"
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
		strtemp = strtemp & "<tr><td height='60' class='tdbg_mainall_lanmu'><font color='#ff9900'>��&nbsp;</font>"
		strtemp = strtemp & "��û���κ���Ŀ�������������Ŀ��"
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
				strtemp = strtemp & "<font color='#ff9900'>��&nbsp;</font>����Ŀû���κ�����"
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
						ClassAD ="<a href='" & rsClassAD("SiteUrl") & "' target='_blank' title='" & rsClassAD("SiteName") & "��" & rsClassAD("SiteUrl") & vbcrlf & rsClassAD("SiteIntro") & "'><img src='" & rsClassAD("ImgUrl") & "'"
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
'���ϸó������ɾ��


'����������(��ȫ���ַ����˵�)
	Public Function Execute(Command)
		If Not IsObject(Conn) Then ConnectionDatabase
		If IsDeBug = 0 Then 
			On Error Resume Next
			Set Execute = Conn.Execute(Command)
			If Err Then
				err.Clear
				Set Conn = Nothing
				Response.Write "��ѯ���ݵ�ʱ���ִ����������Ĳ�ѯ�����Ƿ���ȷ��"
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
		WINNT_CHINESE=(len("����")=2)
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


'��̨����ҳ����ʱ����
Sub dvbbs_error()
	Response.Write"<br>"
	Response.Write"<table cellpadding=3 cellspacing=1 align=center class=""tableBorder"" style=""width:75%"">"
	Response.Write"<tr align=center>"
	Response.Write"<th width=""100%"" height=25 colspan=2>������Ϣ"
	Response.Write"</td>"
	Response.Write"</tr>"
	Response.Write"<tr>"
	Response.Write"<td width=""100%"" class=""Forumrow"" colspan=2>"
	Response.Write ErrMsg
	Response.Write"</td></tr>"
	Response.Write"<tr>"
	Response.Write"<td class=""Forumrow"" valign=middle colspan=2 align=center><a href=""javascript:history.go(-1)""><<������һҳ</a></td></tr>"
	Response.Write"</table>"
	Response.End 
End Sub 

Sub Dv_suc(info)
	Response.Write"<br>"
	Response.Write"<table cellpadding=0 cellspacing=0 align=center class=""tableBorder"" style=""width:75%"">"
	Response.Write"<tr align=center>"
	Response.Write"<th width=""100%"" height=25 colspan=2>�ɹ���Ϣ"
	Response.Write"</td>"
	Response.Write"</tr>"
	Response.Write"<tr>"
	Response.Write"<td width=""100%"" class=""forumRowHighlight"" colspan=2 height=25>"
	Response.Write info
	Response.Write"</td></tr>"
	Response.Write"<tr>"
	Response.Write"<td class=""forumRowHighlight"" valign=middle colspan=2 align=center><a href="&Request.ServerVariables("HTTP_REFERER")&" ><<������һҳ</a></td></tr>"
	Response.Write"</table>"
End Sub

dim UserLogined,UserName,UserLevel,ChargeType,UserPoint,ValidDays
'**************************************************
'��������gotTopic
'��  �ã����ַ���������һ���������ַ���Ӣ����һ���ַ�
'��  ����str   ----ԭ�ַ���
'       strlen ----��ȡ����
'����ֵ����ȡ����ַ���
'**************************************************
function gotTopic(str,strlen)
	if str="" then
		gotTopic=""
		exit function
	end if
	dim l,t,c, i
	str=replace(replace(replace(replace(str,"&nbsp;"," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic=left(str,i) & "��"
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=replace(replace(replace(replace(gotTopic," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function

'**************************************************
'��������JoinChar
'��  �ã����ַ�м��� ? �� &
'��  ����strUrl  ----��ַ
'����ֵ������ ? �� & ����ַ
'**************************************************
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
end function

'==================================================
'��������Announcestr
'��  �ã���ʾ��վ������Ϣ
'��  ����ShowType ------��ʾ��ʽ��1Ϊ����2Ϊ����
'        AnnounceNum  ----�����ʾ����������
'==================================================
function Announcestr(ShowType,AnnounceNum)
	dim sqlAnnounce,rsAnnounce,i,tempAnnouncestr
	if AnnounceNum>0 and AnnounceNum<=10 then
		sqlAnnounce="select top " & AnnounceNum
	else
		sqlAnnounce="select top 10"
	end if
	sqlAnnounce=sqlAnnounce & " * from Announce where IsSelected=True and (ChannelID=0 or ChannelID=" & ChannelID & ") and (ShowType=0 or ShowType=1) order by ID Desc"
	Set rsAnnounce= nt2003.execute(sqlAnnounce)
	if rsAnnounce.bof and rsAnnounce.eof then 
		AnnounceCount=0
		tempAnnouncestr="<p>��ǰû���κι��棡</p>" 
	else 
		AnnounceCount=rsAnnounce.recordcount
		if ShowType=1 then
			do while not rsAnnounce.eof   
				tempAnnouncestr=tempAnnouncestr&"&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' onclick=""javascript:window.open('Announce.asp?ChannelID=" & ChannelID & "&ID=" & rsAnnounce("id") &"', 'newwindow', 'height=300, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='" & rsAnnounce("Content") & "'>" & rsAnnounce("title") & "</div><br><div align='right'>" & rsAnnounce("Author") & "&nbsp;&nbsp;<br>" & FormatDateTime(rsAnnounce("DateAndTime"),1) & "</a>"
				rsAnnounce.movenext
				i=i+1
				if i<AnnounceCount then tempAnnouncestr=tempAnnouncestr& "<hr>"   
			loop
		else
			do while not rsAnnounce.eof   
				tempAnnouncestr=tempAnnouncestr& "&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' onclick=""javascript:window.open('Announce.asp?ChannelID=" & ChannelID & "&ID=" & rsAnnounce("id") &"', 'newwindow', 'height=300, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='" & rsAnnounce("Content") & "' >" & rsAnnounce("title") & "&nbsp;&nbsp;[" & rsAnnounce("Author") & "&nbsp;&nbsp;" & FormatDateTime(rsAnnounce("DateAndTime"),1) & "]</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				rsAnnounce.movenext
			loop
       	end if	
	end if  
	rsAnnounce.close
	set rsAnnounce=nothing
	Announcestr=tempAnnouncestr
end function

'**************************************************
'��������showpage
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sfilename  ----���ӵ�ַ
'       totalnumber ----������
'       maxperpage  ----ÿҳ����
'       ShowTotal   ----�Ƿ���ʾ������
'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת����ĳЩҳ�治��ʹ�ã���������JS����
'       strUnit     ----������λ
'**************************************************
sub showpage(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table align='center'><tr><td>"
	strTemp=strTemp & "�� <font color=blue><b>" & totalnumber & "</b></font> " & strUnit & "&nbsp;&nbsp;&nbsp;"
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "��ҳ ��һҳ&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>��ҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>��һҳ</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "��һҳ βҳ"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>��һҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>βҳ</a>"
  	end if
   	strTemp=strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/ҳ"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;ת����<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">��" & i & "ҳ</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></table>"
	response.write strTemp
end sub


'**************************************************
'��������IsValidEmail
'��  �ã����Email��ַ�Ϸ���
'��  ����email ----Ҫ����Email��ַ
'����ֵ��True  ----Email��ַ�Ϸ�
'       False ----Email��ַ���Ϸ�
'**************************************************
function IsValidEmail(email)
	dim names, name, i, c
	IsValidEmail = true
	names = Split(email, "@")
	if UBound(names) <> 1 then
	   IsValidEmail = false
	   exit function
	end if
	for each name in names
		if Len(name) <= 0 then
			IsValidEmail = false
    		exit function
		end if
		for i = 1 to Len(name)
		    c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
		       IsValidEmail = false
		       exit function
		     end if
	   next
	   if Left(name, 1) = "." or Right(name, 1) = "." then
    	  IsValidEmail = false
	      exit function
	   end if
	next
	if InStr(names(1), ".") <= 0 then
		IsValidEmail = false
	   exit function
	end if
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then
	   IsValidEmail = false
	   exit function
	end if
	if InStr(email, "..") > 0 then
	   IsValidEmail = false
	end if
end function

'**************************************************
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ
'       False ----û�а�װ
'**************************************************
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

'------------------���ĳһĿ¼�Ƿ����-------------------
Function CheckDir(FolderPath)
	dim fso
	folderpath=Server.MapPath(".")&"\"&folderpath
	Set fso1 = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(FolderPath) then
	'����
		CheckDir = True
	Else
	'������
		CheckDir = False
	End if
	Set fso = nothing
End Function

'-------------����ָ����������Ŀ¼---------
Function MakeNewsDir(foldername)
	dim fso,f
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateFolder(foldername)
    MakeNewsDir = True
	Set fso = nothing
End Function


'JMAIL

'**************************************************
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'**************************************************
sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>������Ϣ</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>��������Ŀ���ԭ��</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; ������һҳ</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'**************************************************
'��������WriteSuccessMsg
'��  �ã���ʾ�ɹ���ʾ��Ϣ
'��  ������
'**************************************************
sub WriteSuccessMsg(SuccessMsg)
	dim strSuccess
	strSuccess=strSuccess & "<html><head><title>�ɹ���Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strSuccess=strSuccess & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strSuccess=strSuccess & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center' class='title'><td height='22'><strong>��ϲ�㣡</strong></td></tr>" & vbcrlf
	strSuccess=strSuccess & "  <tr class='tdbg'><td height='100' valign='top'><br>" & SuccessMsg &"</td></tr>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center' class='tdbg'><td>&nbsp;</td></tr>" & vbcrlf
	strSuccess=strSuccess & "</table>" & vbcrlf
	strSuccess=strSuccess & "</body></html>" & vbcrlf
	response.write strSuccess
end sub

'**************************************************
'��������CheckUserLogined
'��  �ã�����û��Ƿ��¼
'��  ������
'����ֵ��True ----�Ѿ���¼
'        False ---û�е�¼
'**************************************************
'delete
'**************************************************
'��������ReplaceBadChar
'��  �ã����˷Ƿ���SQL�ַ�
'��  ����strChar-----Ҫ���˵��ַ�
'����ֵ�����˺���ַ�
'**************************************************
function ReplaceBadChar(strChar)
	if strChar="" then
		ReplaceBadChar=""
	else
		ReplaceBadChar=replace(replace(replace(replace(replace(replace(replace(strChar,"'",""),"*",""),"?",""),"(",""),")",""),"<",""),".","")
	end if
end function

'**************************************************
'��������CheckLevel
'��  �ã�����û�����
'��  ����LevelNum-----Ҫ���ļ���ֵ
'����ֵ����������
'**************************************************
function CheckLevel(LevelNum)
	select case LevelNum
	case 9999
		CheckLevel="�ο�"
	case 999
		CheckLevel="ע���û�"
	case 99
		CheckLevel="�շ��û�"
	case 9
		CheckLevel="VIP�û�"
	case 5
		CheckLevel="����Ա"
	end select
end function

'==================================================
'��������ShowAnnounce
'��  �ã���ʾ��վ������Ϣ
'��  ����ShowType ------��ʾ��ʽ��1Ϊ����2Ϊ����
'        AnnounceNum  ----�����ʾ����������
'==================================================
sub ShowAnnounce(ShowType,AnnounceNum)
	dim sqlAnnounce,rsAnnounce,i
	if AnnounceNum>0 and AnnounceNum<=10 then
		sqlAnnounce="select top " & AnnounceNum
	else
		sqlAnnounce="select top 10"
	end if
	sqlAnnounce=sqlAnnounce & " * from Announce where IsSelected=True and (ChannelID=0 or ChannelID=" & ChannelID & ") and (ShowType=0 or ShowType=1) order by ID Desc"
	Set rsAnnounce= nt2003.execute(sqlAnnounce)
	if rsAnnounce.bof and rsAnnounce.eof then 
		AnnounceCount=0
		response.write "<p>��ǰû���κι��棡</p>" 
	else 
		AnnounceCount=rsAnnounce.recordcount
		if ShowType=1 then
			do while not rsAnnounce.eof   
				response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' onclick=""javascript:window.open('Announce.asp?ChannelID=" & ChannelID & "&ID=" & rsAnnounce("id") &"', 'newwindow', 'height=300, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='" & rsAnnounce("Content") & "'>" & rsAnnounce("title") & "</div><br><div align='right'>" & rsAnnounce("Author") & "&nbsp;&nbsp;<br>" & FormatDateTime(rsAnnounce("DateAndTime"),1) & "</a>"
				rsAnnounce.movenext
				i=i+1
				if i<AnnounceCount then response.write "<hr>"   
			loop
		else
			do while not rsAnnounce.eof   
				response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' onclick=""javascript:window.open('Announce.asp?ChannelID=" & ChannelID & "&ID=" & rsAnnounce("id") &"', 'newwindow', 'height=300, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='" & rsAnnounce("Content") & "' >" & rsAnnounce("title") & "&nbsp;&nbsp;[" & rsAnnounce("Author") & "&nbsp;&nbsp;" & FormatDateTime(rsAnnounce("DateAndTime"),1) & "]</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				rsAnnounce.movenext
			loop
       	end if	
	end if  
	rsAnnounce.close
	set rsAnnounce=nothing
end sub


'==================================================
'��������ShowFriendSite
'��  �ã���ʾ��������վ��
'��  ����LinkType  ----���ӷ�ʽ��1ΪLOGO���ӣ�2Ϊ��������
'       SiteNum   ----�����ʾ���ٸ�վ��
'       Cols      ----�ּ�����ʾ
'       ShowType  ----��ʾ��ʽ��1Ϊ���Ϲ�����2Ϊ�����б�3Ϊ�����б��
'==================================================
Function ShowFriendSite(LinkType,SiteNum,Cols,ShowType)
	dim sqlLink,rsLink,SiteCount,i,strLink
	if LinkType<>1 and LinkType<>2 then
		LinkType=1
	else
		LinkType=Cint(LinkType)
	end if
	if SiteNum<=0 or SiteNum>100 then
		SiteNum=10
	end if
	if Cols<=0 or Cols>20 then
		Cols=10
	end if
	if ShowType=1 then
        	strLink="<div id=rolllink style=overflow:hidden;height:100;width:130><div id=rolllink1>"
	elseif ShowType=3 then
		strLink="<select name='FriendSite' onchange=""if(this.options[this.selectedIndex].value!=''){window.open(this.options[this.selectedIndex].value,'_blank');}""><option value=''>������������վ��</option>"
	end if
	if ShowType=1 or ShowType=2 then
		strLink=strLink & "<table width='100%' cellSpacing='5'><tr align='center' class='tdbg'>"
	end if
	sqlLink="select top " & SiteNum & " * from FriendSite where IsOK=True and LinkType=" & LinkType & " order by IsGood,id desc"
	Set rsLink=Server.CreateObject("ADODB.Recordset")
	rsLink.open sqlLink,conn,1,1
	if rsLink.bof and rsLink.eof then
		if ShowType=1 or ShowType=2 then
	  		for i=1 to SiteNum
				strLink=strLink & "<td><a href='FriendSiteReg.asp' target='_blank'>"
				if LinkType=1 then
					strLink=strLink & "<img src='images/nologo.gif' width='88' height='31' border='0' alt='�������'>"
				else
					strLink=strLink & "�������"
				end if
				strLink=strLink & "</a></td>"
				if i mod Cols=0 and i<SiteNum then
					strLink=strLink & "</tr><tr align='center' class='tdbg'>"
				end if
			next
		end if
	else
		SiteCount=rsLink.recordcount
		for i=1 to SiteCount
			if ShowType=1 or ShowType=2 then
			  if LinkType=1 then
				strLink=strLink & "<td width='88'><a href='" & rsLink("SiteUrl") & "' target='_blank' title='��վ���ƣ�" & rsLink("SiteName") & vbcrlf & "��վ��ַ��" & rsLink("SiteUrl") & vbcrlf & "��վ��飺" & rsLink("SiteIntro") & "'>"
				if rsLink("LogoUrl")="" or rsLink("LogoUrl")="http://" then
					strLink=strLink & "<img src='images/nologo.gif' width='88' height='31' border='0'>"
				else
					strLink=strLink & "<img src='" & rsLink("LogoUrl") & "' width='88' height='31' border='0'>"
				end if
				strLink=strLink & "</a></td>"
			  else
				strLink=strLink & "<td width='88'><a href='" & rsLink("SiteUrl") & "' target='_blank' title='��վ���ƣ�" & rsLink("SiteName") & vbcrlf & "��վ��ַ��" & rsLink("SiteUrl") & vbcrlf & "��վ��飺" & rsLink("SiteIntro") & "'>" & rsLink("SiteName") & "</a></td>"
			  end if
			  if i mod Cols=0 and i<SiteNum then
				strLink=strLink & "</tr><tr align='center' class='tdbg'>"
			  end if
			else
				strLink=strLink & "<option value='" & rsLink("SiteUrl") & "'>" & rsLink("SiteName") & "</option>"
			end if
			rsLink.moveNext
		next
		if SiteCount<SiteNum and (ShowType=1 or ShowType=2) then
			for i=SiteCount+1 to SiteNum
				if LinkType=1 then
					strLink=strLink & "<td width='88'><a href='FriendSiteReg.asp' target='_blank'><img src='images/nologo.gif' width='88' height='31' border='0' alt='�������'></a></td>"
				else
					strLink=strLink & "<td width='88'><a href='FriendSiteReg.asp' target='_blank'>�������</a></td>"
				end if
				if i mod Cols=0 and i<SiteNum then
					strLink=strLink & "</tr><tr align='center' class='tdbg'>"
				end if
			next
		end if
	end if
	if ShowType=1 or ShowType=2 then
		strLink=strLink & "</tr></table>"
	end if
	if ShowType=1 then
	        strLink=strLink & "</div><div id=rolllink2></div></div>"&vbnewline
	elseif ShowType=3 then
		strLink=strLink & "</select>"
	end if
	if ShowType=1 then 
		strLink=strLink & "<script>"&vbnewline
		strLink=strLink & "var rollspeed=40"&vbnewline
		strLink=strLink & "rolllink2.innerHTML=rolllink1.innerHTML"&vbnewline
		strLink=strLink & "function Marquee(){"&vbnewline
		strLink=strLink & "if(rolllink2.offsetTop-rolllink.scrollTop<=0)"&vbnewline
		strLink=strLink & "rolllink.scrollTop-=rolllink1.offsetHeight"&vbnewline
		strLink=strLink & "else{"&vbnewline
		strLink=strLink & "rolllink.scrollTop++"&vbnewline
		strLink=strLink & "}"&vbnewline
		strLink=strLink & "}"&vbnewline
		strLink=strLink & "var MyMar=setInterval(Marquee,rollspeed)"&vbnewline
		strLink=strLink & "rolllink.onmouseover=function() {clearInterval(MyMar)}"&vbnewline
		strLink=strLink & "rolllink.onmouseout=function() {MyMar=setInterval(Marquee,rollspeed)}"&vbnewline
		strLink=strLink & "</script>"&vbnewline
	end if
	rsLink.close
	set rsLink=nothing
	ShowFriendSite=strLink
end Function
jieducm_addrssweb=Request.ServerVariables("server_name") 
if  (jieducm_addrssweb=AdminPurview_class31) or (jieducm_addrssweb=AdminPurview_class32) then
response.Write(jieducm_addrssweb)
response.End()
else

end if
sub ShowGoodSite(SiteNum)
	dim sqlLink,rsLink,SiteCount,i,strLink
	if SiteNum<=0 or SiteNum>100 then
		SiteNum=10
	end if
	strLink=strLink & "<table width='100%' cellSpacing='5'>"
	sqlLink="select top " & SiteNum & " * from FriendSite where IsOK=True and LinkType=1 and IsGood=True order by id desc"
	set rsLink = nt2003.execute(sqlLink)
	if rsLink.bof and rsLink.eof then
	 	for i=1 to SiteNum
			strLink=strLink & "<tr align='center'><td><a href='FriendSiteReg.asp' target='_blank'><img src='images/nologo.gif' width='88' height='31' border='0' alt='�������'></a></td></tr>"
		next
	else
		SiteCount=rsLink.recordcount
		for i=1 to SiteCount
			strLink=strLink & "<tr align='center'><td><a href='" & rsLink("SiteUrl") & "' target='_blank' title='��վ���ƣ�" & rsLink("SiteName") & vbcrlf & "��վ��ַ��" & rsLink("SiteUrl") & vbcrlf & "��վ��飺" & rsLink("SiteIntro") & "'>"
			if rsLink("LogoUrl")="" or rsLink("LogoUrl")="http://" then
				strLink=strLink & "<img src='images/nologo.gif' width='88' height='31' border='0'>"
			else
				strLink=strLink & "<img src='" & rsLink("LogoUrl") & "' width='88' height='31' border='0'>"
			end if
			strLink=strLink & "</a></td></tr>"
			rsLink.moveNext
		next
		for i=SiteCount+1 to SiteNum
			strLink=strLink & "<tr align='center'><td><a href='FriendSiteReg.asp' target='_blank'><img src='images/nologo.gif' width='88' height='31' border='0' alt='�������'></a></td></tr>"
		next
	end if
	strLink=strLink & "</table>"
	response.write strLink
	rsLink.close
	set rsLink=nothing

end sub

'==================================================
'��������ShowUserLogin
'��  �ã���ʾ�û���¼��
'��  ������
'==================================================
'DELETE

'==================================================
'��������ShowTopUser
'��  �ã���ʾ�û����У����ѷ������������������ȣ��ٰ�ע���Ⱥ�˳������
'��  ����UserNum-------��ʾ���û�����
'==================================================
'DELETE
'==================================================
'��������ShowAllUser
'��  �ã���ҳ��ʾ�����û�
'��  ������
'==================================================
'DELETE
'==================================================
'��������PopAnnouceWindow
'��  �ã��������洰��
'��  ����Width-------�������ڿ��
'		 Height------�������ڸ߶�
'==================================================
sub PopAnnouceWindow(Width,Height)
	dim popCount,rsAnnounce
	set rsAnnounce=nt2003.execute("select count(*) from Announce where IsSelected=True and (ChannelID=0 or ChannelID=" & ChannelID & ") and (ShowType=0 or ShowType=2)")
	popCount=rsAnnounce(0)
	if popCount>0 then
		if  nt2003.site_setting(13)=1 and session("Poped")<>ChannelID then
			response.write "<script LANGUAGE='JavaScript'>"
			response.write "window.open ('Announce.asp?ChannelID=" & ChannelID & "', 'newwindow', 'height=" & Height & ", width=" & Width & ", toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"
			response.write "</script>"
			session("Poped")=ChannelID
		end if
	end if
end sub

'==================================================
'��������ShowPath
'��  �ã���ʾ������������λ�á�������Ϣ
'��  ������
'==================================================
sub ShowPath()
	if PageTitle<>"" and ChannelID<>1 then
		strPath=strPath & "&nbsp;&gt;&gt;&nbsp;" & PageTitle
	end if
	response.write strPath
end sub

'==================================================
'��������MenuJS
'��  �ã���ʾ���ε�����JS����
'��  ������
'==================================================
sub MenuJS()
  if nt2003.site_setting(4)=1 then
%>
<script language="JavaScript" type="text/JavaScript">
//���ε�����JS����
document.write("<style type=text/css>#master {LEFT: -200px; POSITION: absolute; TOP: 25px; VISIBILITY: visible; Z-INDEX: 999}</style>")
document.write("<table id=master width='218' border='0' cellspacing='0' cellpadding='0'><tr><td><img border=0 height=6 src=images/menutop.gif  width=200></td><td rowspan='2' valign='top'><img id=menu onMouseOver=javascript:expand() border=0 height=70 name=menutop src=images/menuo.gif width=18></td></tr>");
document.write("<tr><td valign='top'><table width='100%' border='0' cellspacing='5' cellpadding='0'><tr><td height='400' valign='top'><table width=100% height='100%' border=1 cellpadding=0 cellspacing=5 bordercolor='#666666' bgcolor=#ecf6f5 style=FILTER: alpha(opacity=90)><tr>");
document.write("<td height='10' align='center' bordercolor='#ecf6f5'><font color=999900><strong>�� Ŀ �� �� �� ��</strong></font></td></tr><tr><td valign='top' bordercolor='#ecf6f5'>");
document.write("<iframe width=100% height=100% src='classtree.asp?ChannelID=<%=ChannelID%>' frameborder=0></iframe></td></tr></table></td></tr></table></td></tr></table>");

var ie = document.all ? 1 : 0
var ns = document.layers ? 1 : 0
var master = new Object("element")
master.curLeft = -200;	master.curTop = 10;
master.gapLeft = 0;		master.gapTop = 0;
master.timer = null;

if(ie){var sidemenu = document.all.master;}
if(ns){var sidemenu = document.master;}
setInterval("FixY()",100);

function moveAlong(layerName, paceLeft, paceTop, fromLeft, fromTop){
	clearTimeout(eval(layerName).timer)
	if(eval(layerName).curLeft != fromLeft){
		if((Math.max(eval(layerName).curLeft, fromLeft) - Math.min(eval(layerName).curLeft, fromLeft)) < paceLeft){eval(layerName).curLeft = fromLeft}
		else if(eval(layerName).curLeft < fromLeft){eval(layerName).curLeft = eval(layerName).curLeft + paceLeft}
			else if(eval(layerName).curLeft > fromLeft){eval(layerName).curLeft = eval(layerName).curLeft - paceLeft}
		if(ie){document.all[layerName].style.left = eval(layerName).curLeft}
		if(ns){document[layerName].left = eval(layerName).curLeft}
	}
	if(eval(layerName).curTop != fromTop){
   if((Math.max(eval(layerName).curTop, fromTop) - Math.min(eval(layerName).curTop, fromTop)) < paceTop){eval(layerName).curTop = fromTop}
		else if(eval(layerName).curTop < fromTop){eval(layerName).curTop = eval(layerName).curTop + paceTop}
			else if(eval(layerName).curTop > fromTop){eval(layerName).curTop = eval(layerName).curTop - paceTop}
		if(ie){document.all[layerName].style.top = eval(layerName).curTop}
		if(ns){document[layerName].top = eval(layerName).curTop}
	}
	eval(layerName).timer=setTimeout('moveAlong("'+layerName+'",'+paceLeft+','+paceTop+','+fromLeft+','+fromTop+')',30)
}

function setPace(layerName, fromLeft, fromTop, motionSpeed){
	eval(layerName).gapLeft = (Math.max(eval(layerName).curLeft, fromLeft) - Math.min(eval(layerName).curLeft, fromLeft))/motionSpeed
	eval(layerName).gapTop = (Math.max(eval(layerName).curTop, fromTop) - Math.min(eval(layerName).curTop, fromTop))/motionSpeed
	moveAlong(layerName, eval(layerName).gapLeft, eval(layerName).gapTop, fromLeft, fromTop)
}

var expandState = 0

function expand(){
	if(expandState == 0){setPace("master", 0, 10, 10); if(ie){document.menutop.src = "images/menui.gif"}; expandState = 1;}
	else{setPace("master", -200, 10, 10); if(ie){document.menutop.src = "images/menuo.gif"}; expandState = 0;}
}

function FixY(){
	if(ie){sidemenu.style.top = document.body.scrollTop+10}
	if(ns){sidemenu.top = window.pageYOffset+10}
}
</script>
<%
end if
end sub
sub total_jieducm()
response.Write("<div style='DISPLAY: none'>")
response.Write("<script language=javascript type=text/javascript src=http://js.users.51.la/3906722.js></script>")
response.Write("</div>")
end sub
'==================================================
'��������ShowSearchForm
'��  �ã���ʾ����������
'��  ����ShowType ----��ʾ��ʽ��1Ϊ���ģʽ��2Ϊ��׼ģʽ��3Ϊ�߼�ģʽ
'==================================================
sub ShowSearchForm(Action,ShowType)
	if ShowType<>1 and ShowType<>2 and ShowType<>3 then
		ShowType=1
	end if
	response.write "<table border='0' cellpadding='0' cellspacing='0'>"
	response.write "<form method='Get' name='SearchForm' action='" & Action & "'>"
	response.write "<tr><td height='28' align='center'>"
	if ShowType=1 then
		response.write "<input type='text' name='keyword'  size='15' value='�ؼ���' maxlength='50' onFocus='this.select();'>&nbsp;"
		response.write "<input type='hidden' name='field' value='Title'>"
		response.write "<input type='submit' name='Submit'  value='��&nbsp;��'>" 
	elseif Showtype=2 then
		response.write "<select name='Field' size='1'>"
    	if ChannelID=2 then
			response.write "<option value='Title' selected>���±���</option>"
			response.write "<option value='Content'>��������</option>"
			response.write "<option value='Author'>��������</option>"
			response.write "<option value='Editor'>�༭����</option>"
		elseif ChannelID=3 then	
			response.write "<option value='SoftName' selected>�������</option>"
			response.write "<option value='SoftIntro'>������</option>"
			response.write "<option value='Author'>�������</option>"
			response.write "<option value='Editor'>�༭����</option>"
		elseif ChannelID=4 then	
			response.write "<option value='PhotoName' selected>ͼƬ����</option>"
			response.write "<option value='PhotoIntro'>ͼƬ���</option>"
			response.write "<option value='Author'>ͼƬ����</option>"
			response.write "<option value='Editor'>�༭����</option>"
		else
			response.write "<option value='Title' selected>���±���</option>"
			response.write "<option value='Content'>��������</option>"
			response.write "<option value='Author'>��������</option>"
			response.write "<option value='Editor'>�༭����</option>"
		end if
		response.write "</select>&nbsp;"
		response.write "<select name='ClassID'><option value=''>������Ŀ</option>"
		call Admin_ShowClass_Option(5,0)
		response.write "</select>&nbsp;<input type='text' name='keyword'  size='20' value='�ؼ���' maxlength='50' onFocus='this.select();'>&nbsp;"
		response.write "<input type='submit' name='Submit'  value='��&nbsp;��'>" 
	elseif Showtype=3 then
	
	end if
	response.write "</td></tr></form></table>"
end sub


'==================================================
'��������FloatAD
'��  �ã��������
'��  ������
'==================================================
sub FloatAD()
%>
<SCRIPT language=javascript>
<!--moving logo-->
window.onload=FlAD;
var brOK=false;
var mie=false;
var aver=parseInt(navigator.appVersion.substring(0,1));
var aname=navigator.appName;
var mystop=0;

function checkbrOK()
{if(aname.indexOf("Internet Explorer")!=-1)
{if(aver>=4) brOK=navigator.javaEnabled();
mie=true;
}
if(aname.indexOf("Netscape")!=-1)  
{if(aver>=4) brOK=navigator.javaEnabled();}
}
var vmin=2;
var vmax=5;
var vr=2;
var timer1;

function Chip(chipname,width,height)
{this.named=chipname;
this.vx=vmin+vmax*Math.random();
this.vy=vmin+vmax*Math.random();
this.w=width;
this.h=height;
this.xx=0;
this.yy=0;
this.timer1=null;
}

function movechip(chipname)
{
if(brOK && mystop==0)
{eval("chip="+chipname);
if(!mie)
{pageX=window.pageXOffset;
pageW=window.innerWidth;
pageY=window.pageYOffset;
pageH=window.innerHeight;
}
else
{pageX=window.document.body.scrollLeft;
pageW=window.document.body.offsetWidth-8;
pageY=window.document.body.scrollTop;
pageH=window.document.body.offsetHeight;
} 
chip.xx=chip.xx+chip.vx;
chip.yy=chip.yy+chip.vy;
chip.vx+=vr*(Math.random()-0.5);
chip.vy+=vr*(Math.random()-0.5);
if(chip.vx>(vmax+vmin))  chip.vx=(vmax+vmin)*2-chip.vx;
if(chip.vx<(-vmax-vmin)) chip.vx=(-vmax-vmin)*2-chip.vx;
if(chip.vy>(vmax+vmin))  chip.vy=(vmax+vmin)*2-chip.vy;
if(chip.vy<(-vmax-vmin)) chip.vy=(-vmax-vmin)*2-chip.vy;
if(chip.xx<=pageX)
{chip.xx=pageX;
chip.vx=vmin+vmax*Math.random();
}
if(chip.xx>=pageX+pageW-chip.w)
{chip.xx=pageX+pageW-chip.w;
chip.vx=-vmin-vmax*Math.random();
}
if(chip.yy<=pageY)
{chip.yy=pageY;
chip.vy=vmin+vmax*Math.random();
}
if(chip.yy>=pageY+pageH-chip.h)
{chip.yy=pageY+pageH-chip.h;
chip.vy=-vmin-vmax*Math.random();
}
if(!mie)
{eval('document.'+chip.named+'.top ='+chip.yy);
eval('document.'+chip.named+'.left='+chip.xx);
} 
else
{eval('document.all.'+chip.named+'.style.pixelLeft='+chip.xx);
eval('document.all.'+chip.named+'.style.pixelTop ='+chip.yy); 
}
	chip.timer1=setTimeout("movechip('"+chip.named+"')",100);
}
}
function stopme(x)
{
brOk=true;
mystop=x;
movechip("FlAD");
}
var FlAD;
var chip;
function FlAD()
{checkbrOK(); 
FlAD=new Chip("FlAD",80,80);
if(brOK) 
{ movechip("FlAD");
}
}
ns4=(document.layers)?true:false;
ie4=(document.all)?true:false;

function cncover()
{
if(ns4){
	//document.cnc.left=window.innerWidth/2-400;
	document.FlAD.visibility="hide";
	eval('document.cnc.left=document.'+chip.named+'.left');
	eval('document.cnc.top=document.'+chip.named+'.top-15');	
	document.cnc.visibility="show";
	}else if(ie4) 
	{
	document.all.FlAD.style.visibility="hidden";
	//document.all.cnc.style.left=window.document.body.offsetWidth/2-400;
	document.all.cnc.style.left=parseInt(document.all.FlAD.style.left)-0;
	document.all.cnc.style.top=parseInt(document.all.FlAD.style.top)-0;	
	document.all.cnc.style.visibility="visible";
	stopme(1);
	}
}

function cncout()
{
if(ns4){
	document.cnc.visibility="hide";
	document.FlAD.visibility="show";
	}else if(ie4) 
	{
	document.all.cnc.style.visibility="hidden";
	document.all.FlAD.style.visibility="visible";
	stopme(0);
	}
}
</script>
<%
end sub


'==================================================
'��������FixedAD
'��  �ã��̶�λ�ù��
'��  ������
'==================================================
sub FixedAD()
%>
<script LANGUAGE="JavaScript">
<!-- Begin
var imgheight
var imgleft
document.ns = navigator.appName == "Netscape"
if (navigator.appName == "Netscape")
{
imgheight=document.FixAD.pageY
imgleft=document.FixAD.pageX
}
else
{
imgheight=600-parseInt(FixAD.style.top)
imgleft=parseInt(FixAD.style.left)
}
myload()
function myload()
{
if (navigator.appName == "Netscape")
{document.FixAD.pageY=pageYOffset+window.innerHeight-imgheight;
document.FixAD.pageX=imgleft;
leftmove();
}
else
{
FixAD.style.top=document.body.scrollTop+document.body.offsetHeight-imgheight;
FixAD.style.left=imgleft;
leftmove();
}
}
function leftmove()
 {
 if(document.ns)
 {
 document.FixAD.top=pageYOffset+window.innerHeight-imgheight
 document.FixAD.left=imgleft;
 setTimeout("leftmove();",50)
 }
 else
 {
 FixAD.style.top=document.body.scrollTop+document.body.offsetHeight-imgheight;
 FixAD.style.left=imgleft;
 setTimeout("leftmove();",50)
 }
 }

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true)
//  End -->
</script>
<%
end sub

'==================================================
'��������FixedAD1
'��  �ã��̶�λ�ù�棨ͼƬλ�ó�������ʱ��ʱ�����⣩
'��  ������
'==================================================
sub FixedAD1()
%>
<script LANGUAGE="JavaScript">
<!-- Begin
        self.onError=null;
        currentX = currentY = 0;
        whichIt = null;
        lastScrollX = 0; lastScrollY = 0;
        NS = (document.layers) ? 1 : 0;
        IE = (document.all) ? 1: 0;
        function heartBeat() {
                if(IE) { diffY = document.body.scrollTop; diffX = document.body.scrollLeft; }
            if(NS) { diffY = self.pageYOffset; diffX = self.pageXOffset; }
                if(diffY != lastScrollY) {
                        percent = .1 * (diffY - lastScrollY);
                        if(percent > 0) percent = Math.ceil(percent);
                        else percent = Math.floor(percent);
                                        if(IE) document.all.FixAD.style.pixelTop += percent;
                                        if(NS) document.FixAD.top += percent;
                        lastScrollY = lastScrollY + percent;
            }
                if(diffX != lastScrollX) {
                        percent = .1 * (diffX - lastScrollX);
                        if(percent > 0) percent = Math.ceil(percent);
                        else percent = Math.floor(percent);
                        if(IE) document.all.FixAD.style.pixelLeft += percent;
                        if(NS) document.FixAD.left += percent;
                        lastScrollX = lastScrollX + percent;
                }
        }
        function checkFocus(x,y) {
                stalkerx = document.FixAD.pageX;
                stalkery = document.FixAD.pageY;
                stalkerwidth = document.FixAD.clip.width;
                stalkerheight = document.FixAD.clip.height;
                if( (x > stalkerx && x < (stalkerx+stalkerwidth)) && (y > stalkery && y < (stalkery+stalkerheight))) return true;
                else return false;
        }
        function grabIt(e) {
                if(IE) {
                        whichIt = event.srcElement;
                        while (whichIt.id.indexOf("FixAD") == -1) {
                                whichIt = whichIt.parentElement;
                                if (whichIt == null) { return true; }
                    }
                        whichIt.style.pixelLeft = whichIt.offsetLeft;
                    whichIt.style.pixelTop = whichIt.offsetTop;
                        currentX = (event.clientX + document.body.scrollLeft);
                           currentY = (event.clientY + document.body.scrollTop);
                } else {
                window.captureEvents(Event.MOUSEMOVE);
                if(checkFocus (e.pageX,e.pageY)) {
                        whichIt = document.FixAD;
                        StalkerTouchedX = e.pageX-document.FixAD.pageX;
                        StalkerTouchedY = e.pageY-document.FixAD.pageY;
                }
                }
            return true;
        }
        function moveIt(e) {
                if (whichIt == null) { return false; }
                if(IE) {
                    newX = (event.clientX + document.body.scrollLeft);
                    newY = (event.clientY + document.body.scrollTop);
                    distanceX = (newX - currentX);    distanceY = (newY - currentY);
                    currentX = newX;    currentY = newY;
                    whichIt.style.pixelLeft += distanceX;
                    whichIt.style.pixelTop += distanceY;
                        if(whichIt.style.pixelTop < document.body.scrollTop) whichIt.style.pixelTop = document.body.scrollTop;
                        if(whichIt.style.pixelLeft < document.body.scrollLeft) whichIt.style.pixelLeft = document.body.scrollLeft;
                        if(whichIt.style.pixelLeft > document.body.offsetWidth - document.body.scrollLeft - whichIt.style.pixelWidth - 20) whichIt.style.pixelLeft = document.body.offsetWidth - whichIt.style.pixelWidth - 20;
                        if(whichIt.style.pixelTop > document.body.offsetHeight + document.body.scrollTop - whichIt.style.pixelHeight - 5) whichIt.style.pixelTop = document.body.offsetHeight + document.body.scrollTop - whichIt.style.pixelHeight - 5;
                        event.returnValue = false;
                } else {
                        whichIt.moveTo(e.pageX-StalkerTouchedX,e.pageY-StalkerTouchedY);
                if(whichIt.left < 0+self.pageXOffset) whichIt.left = 0+self.pageXOffset;
                if(whichIt.top < 0+self.pageYOffset) whichIt.top = 0+self.pageYOffset;
                if( (whichIt.left + whichIt.clip.width) >= (window.innerWidth+self.pageXOffset-17)) whichIt.left = ((window.innerWidth+self.pageXOffset)-whichIt.clip.width)-17;
                if( (whichIt.top + whichIt.clip.height) >= (window.innerHeight+self.pageYOffset-17)) whichIt.top = ((window.innerHeight+self.pageYOffset)-whichIt.clip.height)-17;
                return false;
                }
            return false;
        }
        function dropIt() {
                whichIt = null;
            if(NS) window.releaseEvents (Event.MOUSEMOVE);
            return true;
        }
        if(NS) {
                window.captureEvents(Event.MOUSEUP|Event.MOUSEDOWN);
                window.onmousedown = grabIt;
                 window.onmousemove = moveIt;
                window.onmouseup = dropIt;
        }
        if(IE) {
                document.onmousedown = grabIt;
                 document.onmousemove = moveIt;
                document.onmouseup = dropIt;
        }
        if(NS || IE) action = window.setInterval("heartBeat()",1);
//  End -->
</script>
<%
end sub

sub ShowAD(ADType)
	dim sqlAD,rsAD,AD,arrSetting,popleft,poptop,floatleft,floattop,fixedleft,fixedtop
	sqlAD="select * from Advertisement where IsSelected=True"
	sqlAD=sqlAD & " and (ChannelID=0 or ChannelID=" & ChannelID & ")"
	sqlAD=sqlAD & " and ADType=" & ADtype & " order by ID Desc"
	set rsAD=nt2003.execute(sqlAD)
	if not rsAd.bof and not rsAD.eof then
		do while not rsAD.eof
			if rsAD("isflash")=true then
				AD= "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0'"
				if rsAD("ImgWidth")>0 then AD = AD & " width='" & rsAD("ImgWidth") & "'"
				if rsAD("ImgHeight")>0 then AD = AD & " height='" & rsAD("ImgHeight") & "'"
				AD = AD & " id=OBJECT1><param name='movie' value='" & rsAD("ImgUrl") & "'><param name='quality' value='high'><embed src='" & rsAD("ImgUrl") & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'"
				if rsAD("ImgWidth")>0 then AD = AD & " width='" & rsAD("ImgWidth") & "'"
				if rsAD("ImgHeight")>0 then AD = AD & " height='" & rsAD("ImgHeight") & "'"
				AD = AD & "></embed></object>"
			else
				AD ="<a href='" & rsAD("SiteUrl") & "' target='_blank' title='" & rsAD("SiteName") & "��" & rsAD("SiteUrl") & "'><img src='" & rsAD("ImgUrl") & "'"
				if rsAD("ImgWidth")>0 then AD = AD & " width='" & rsAD("ImgWidth") & "'"
				if rsAD("ImgHeight")>0 then AD = AD & " height='" & rsAD("ImgHeight") & "'"
				AD = AD & " border='0'></a>"
			end if
			if ADtype=0 then
				if  session("PopAD"&rsAD("ID")&ChannelID)<>True then
					if instr(rsAD("ADSetting"),"|")>0 then
						arrSetting=split(rsAD("ADSetting"),"|")
						popleft=arrsetting(0)
						poptop=arrsetting(1)
					end if
					response.write "<SCRIPT language=javascript>"
					response.write "window.open(""PopAD.asp?Id="& rsAD("ID")&""",""popad"&rsAD("ID")&""",""toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,width="&rsAD("ImgWidth")&",height="&rsAD("ImgHeight")&",top="&poptop&",left="&popleft&""");"
					response.write "</SCRIPT>"
					session("PopAD"&rsAD("ID")&ChannelID)=True
				end if
			elseif ADtype=1 then
				response.write AD
				exit do
			elseif ADtype=2 then
				response.write AD
				exit do
			elseif ADtype=3 then
				response.write AD
				exit do
			elseif ADtype=4 then
				if instr(rsAD("ADSetting"),"|")>0 then
					arrSetting=split(rsAD("ADSetting"),"|")
					floatleft=arrsetting(0)
					floattop=arrsetting(1)
				end if
				response.write "<div id='FlAD' style='position:absolute; z-index:10;left: "&floatleft&"; top: "&floattop&"'>" & AD & "</div>"
				call FloatAD()
				exit do
			elseif ADtype=5 then
				if instr(rsAD("ADSetting"),"|")>0 then
					arrSetting=split(rsAD("ADSetting"),"|")
					fixedleft=arrsetting(0)
					fixedtop=arrsetting(1)
				end if
				response.write "<div id='FixAD' style='position:absolute; z-index:10;left: "&fixedleft&"; top: "&fixedtop&"'>" & AD & "</div>"
				call FixedAD()
				exit do
			end if
			rsAD.movenext
		loop
	end if
	rsAD.close
	set rsAD=nothing
end sub
%>
