<%@language=vbscript codepage=936 %>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/md5.asp"-->

<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷����ݶȴ�ý
'�ٷ���ҳ��http://www.jieducm.cn
'����֧�֣�robin 
'QQ;616591415
'�绰��13598006137
'��Ȩ����Ȩ���ڽݶȴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2008��2011 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
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
	ErrMsg=ErrMsg & "<br><li>�û�������Ϊ�գ�</li>"
end if
if Password="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>���벻��Ϊ�գ�</li>"
end if
if CheckCode="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>��֤�벻��Ϊ�գ�</li>"
end if
if session("CheckCode")="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>���¼ʱ������������·��ص�¼ҳ����е�¼��</li>"
end if
if CheckCode<>CStr(session("CheckCode")) then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�������ȷ�����ϵͳ�����Ĳ�һ�£����������롣</li>"
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
		rs_jieducm("class")="��¼��̨"
		rs_jieducm("content")=username&"��¼ʱ��:"&now&"��¼IP:"&Request.ServerVariables("REMOTE_ADDR")&""
		rs_jieducm("jiegou")="ʧ��"
    	rs_jieducm.update
    	rs_jieducm.close
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�û�����������󣡣����㱾�ε�½��ip:"&Request.ServerVariables("REMOTE_ADDR")&"�Ѽ�¼�ڰ�</li>"
		
	else
		if password<>rs("password") then
		Set rs_jieducm=server.createobject("ADODB.RECORDSET")
	    rs_jieducm.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rs_jieducm.addnew
		rs_jieducm("username")=username
		rs_jieducm("class")="��¼��̨"
		rs_jieducm("content")=username&"��¼ʱ��:"&now&"��¼IP:"&Request.ServerVariables("REMOTE_ADDR")&""
		rs_jieducm("jiegou")="ʧ��"
    	rs_jieducm.update
    	rs_jieducm.close
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>�û�����������󣡣��㱾�ε�½��ip:"&Request.ServerVariables("REMOTE_ADDR")&"�Ѽ�¼�ڰ���</li>"
		else
		 
			nt2003.execute("update "&jieducm&"_Admin set LastLoginIP='"&Request.ServerVariables("REMOTE_ADDR")&"',LastLoginTime='"&now&"',LoginTimes=LoginTimes+1 Where username='"&username&"'") 
			session.Timeout=nt2003.site_setting(15)
			session("AdminName")=username
		
		Set rs_jieducm=server.createobject("ADODB.RECORDSET")
	    rs_jieducm.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rs_jieducm.addnew
		rs_jieducm("username")=username
		rs_jieducm("class")="��¼��̨"
		rs_jieducm("content")=username&"��¼ʱ��:"&now&"��¼IP:"&Request.ServerVariables("REMOTE_ADDR")&""
		rs_jieducm("jiegou")="�ɹ�"
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
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'****************************************************
sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td height='22' class='title'><strong>������Ϣ</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr><td height='100' class='tdbg' valign='top'><b>��������Ŀ���ԭ��</b><br>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td class='tdbg'><a href='Admin_Login.asp'>&lt;&lt; ���ص�¼ҳ��</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub
%>