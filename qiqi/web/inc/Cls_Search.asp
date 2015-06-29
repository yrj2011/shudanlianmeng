<%

Class Cls_Search
	Public sKeyWord						'关键字
	Public Inc_Url						'链接
	Public Inc_SQL						'组装SQL
	Private Inc_Tmp						'临时
	Private sTypes						'搜索类型
	Private sTerm						'
	'//* 类初始化 *//
	Private Sub Class_Initialize()
		sKeyWord=""
		sTypes=""
		sTerm=""
		Inc_Url="?"
		Inc_Tmp=""
	End Sub

	Public Property Get KeyWord
		KeyWord=sKeyWord
	End Property

	Public Sub Format_Search(stype_1,fsTypes,sqlt,st,vt)
		sKeyWord=YSvoid.Code_Query("keyword")
		If sKeyWord<>"" Then
			sTypes=YSvoid.Code_Query("sea_type")
			sTerm=YSvoid.Code_Query("sea_term")
			If InStr(","&fsTypes&",",","&sTypes&",")<1 then sTypes=stype_1
			if sTerm<>"all" then sTerm="only"
			Inc_Tmp=sql_key(keyword,sTypes,sqlt,sTerm)
			If Inc_Tmp<>"" Then Inc_Url=Inc_Url&"sea_type="&sTypes&"&sea_term="&sTerm&"&keyword="&server.urlencode(sKeyWord)&"&"
		End If
		if st=1 and Inc_Tmp<>"" then
			if Inc_SQL="" then
				'Inc_SQL=" where "&Inc_Tmp
				Inc_SQL=Inc_Tmp
			else
				Inc_SQL=Inc_SQL&" and "&Inc_Tmp
			End If
		End If
	End Sub

	Public Function SQL_Key(keyes,fields,kt,kterm)
		If YSvoid.Is_Null(keyes)="" Or YSvoid.Is_Null(fields)="" Then
			SQL_Key=""
			Exit Function
		End If
		dim ddim,dnum,di,spvar,termvar
		spvar=","
		termvar="or"
		if kt=2 then spvar=" "
		if kterm="all" then termvar="and"
		keyes=replace(keyes,"'","")
		ddim=split(keyes,spvar)
		dnum=ubound(ddim)
		SQL_Key="("&fields&" like '%"&ddim(0)&"%'"
		for di=1 to dnum
			SQL_Key=SQL_Key&" "&termvar&" "&fields&" like '%"&ddim(di)&"%'"
		next
		erase ddim
		SQL_Key=SQL_Key&")"
	End Function
End Class

%>