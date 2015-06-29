<%

'* ----------------------------------------
'* 文件说明
'* 作  用：用于前、后台频道分类基本信息处理
'* ----------------------------------------

Dim ChannelSort			'频道识别码，一般与字段名相同
Dim Cid					'频道分类ID
Dim ClassTrue			'判断分类是否存在
Dim ClassArr			'分类数组
Dim ClassNum			'分类数组个数
Dim ClassColNum			'余数

Cid = 0
ClassTrue = True
ClassNum = -1
ClassColNum = 7

'//* 分类数据初始化 *//
Sub Class_Load()
	YSvoid.Name = "Class_"&ChannelID
	If YSvoid.Cache_Chk() Then
		Call Class_LoadData()
	End If
	ClassArr = YSvoid.Value
	If Not IsArray(ClassArr) Then
		ClassTrue = False
		Exit Sub
	End If
	ClassNum = Ubound(ClassArr,2)
End Sub

'//* 分类数据读取 *//
Sub Class_LoadData()
    SQL = "Select C_ID,ClassName,ClassParent,ClassParentStr,ClassDepth,ClassChild,ClassID,ClassChildStr From YSvoid_Class Where ChannelID="&ChannelID&" And ClassDepth<"&YSvoid.ClassDepth&" Order By ClassID,ClassOrder,ClassDepth"
	Set Rs = YSvoid.Exec(SQL,1)
	If Rs.Eof Then
		Rs.Close
		YSvoid.Value = "kong"
		Exit Sub
	End If
	YSvoid.Value = Rs.GetRows()
	Rs.Close
	Exit Sub
End Sub

Function Class_Grade(vCid,vDepth)
	Dim Tmpstr, vClassParentStr, CurrenCid, Vii
	If vCid = 0 Then Exit Function
	'Tmpstr = ChannelName
	For Vii = 0 To ClassNum
		If Int(ClassArr(0,Vii)) = Int(vCid) Then
			vClassParentStr = ClassArr(3,Vii)
			Exit For
		End If
	Next
	Vii = 0
	If Instr(vClassParentStr,",")<>0 Then
		vClassParentStr = Split(vClassParentStr,",")
		For Vii = 0 To Ubound(vClassParentStr)
			CurrenCid = vClassParentStr(Vii)	
			    Tmpstr = Tmpstr & "&nbsp;<a href='List.asp?c_id="&CurrenCid&"' class=h_menu>"&Class_Name_Get(CurrenCid)&"</a>&nbsp;→&nbsp;"	
			If Vii=vDepth-1 Then Exit For
		Next
		Erase vClassParentStr
	Else	
			  Tmpstr = Tmpstr & "&nbsp;<a href='List.asp?c_id="&vCid&"' class=h_menu>"&Class_Name_Get(vCid)&"</a>&nbsp;→&nbsp;"
	End If
	Class_Grade = Tmpstr
End Function

Function Class_Grade_List(vCid,vDepth,vClass)
	Dim Tmpstr, vClassParentStr, CurrenCid, Vii
	If vCid = 0 Then Exit Function
	'Tmpstr = ChannelName
	For Vii = 0 To ClassNum
		If Int(ClassArr(0,Vii)) = Int(vCid) Then
			vClassParentStr = ClassArr(3,Vii)
			Exit For
		End If
	Next
	Vii = 0
	If Instr(vClassParentStr,",")<>0 Then
		vClassParentStr = Split(vClassParentStr,",")
		For Vii = 0 To Ubound(vClassParentStr)
			CurrenCid = vClassParentStr(Vii)	
			    Tmpstr = Tmpstr & "&nbsp;<a href='List.asp?c_id="&CurrenCid&"' class="&vClass&">"&Class_Name_Get(CurrenCid)&"</a>&nbsp;→&nbsp;"	
			If Vii=vDepth-1 Then Exit For
		Next
		Erase vClassParentStr
	Else	
			  Tmpstr = Tmpstr & "&nbsp;<a href='List.asp?c_id="&vCid&"' class="&vClass&">"&Class_Name_Get(vCid)&"</a>&nbsp;→&nbsp;"
	End If
	Class_Grade_List = Tmpstr
End Function

Function Class_Grade_Xml(vCid,vDepth)
	Dim Tmpstr, vClassParentStr, CurrenCid, Vii
	If vCid = 0 Then Exit Function
	'Tmpstr = ChannelName
	For Vii = 0 To ClassNum
		If Int(ClassArr(0,Vii)) = Int(vCid) Then
			vClassParentStr = ClassArr(3,Vii)
			Exit For
		End If
	Next
	Vii = 0
	If Instr(vClassParentStr,",")<>0 Then
		vClassParentStr = Split(vClassParentStr,",")
		For Vii = 0 To Ubound(vClassParentStr)
			CurrenCid = vClassParentStr(Vii)	
			    Tmpstr = Tmpstr &Class_Name_Get(CurrenCid)&" - "
			If Vii=vDepth-1 Then Exit For
		Next
		Erase vClassParentStr
	Else	
			  Tmpstr = Tmpstr & Class_Name_Get(vCid)&" - "
	End If
	Class_Grade_Xml = Tmpstr
End Function

'//* 当前所在分类名称 *//
Function Class_Name_Get(vCid)
	Dim Ni
	Class_Name_Get=""
	If Not ClassTrue Then Exit Function
	if Int(vCid)=0 Then Exit Function
	For Ni=0 To ClassNum
		If Int(ClassArr(0,Ni))=Int(vCid) Then
			Class_Name_Get=ClassArr(1,Ni)
			Exit For
		End If
	Next
End Function

'//* 当前所在分类的子分类数量 *//
Function Class_Parent_Get(vCid)
	Dim Ni
	Class_Parent_Get=0
	If Not ClassTrue Then Exit Function
	if Int(vCid)=0 Then Exit Function
	For Ni=0 To ClassNum
		If Int(ClassArr(0,Ni))=Int(vCid) Then
			Class_Parent_Get=ClassArr(5,Ni)
			Exit For
		End If
	Next
End Function
'//* 当前分类及子分类 *//
Function Class_Cid_Get(vCid)
	Dim Ni, TmpStr, nCid, ClassParent
	nCid=0
	Class_Cid_Get=""
	If Not ClassTrue Then Exit Function
	If Int(vCid)=0 Then Exit Function
	If Int(vCid)=0 Then TmpStr = ""

	For Ni=0 To ClassNum
		If Int(ClassArr(0,Ni))=Int(vCid) Then
			Class_Cid_Get=ClassArr(7,Ni)
			Exit For
		End If
	Next
End Function

Function Class_Cid_Get2(vCid)
	Dim Ni, TmpStr, nCid, ClassParent
	nCid=0
	Class_Cid_Get2=""
	If Not ClassTrue Then Exit Function
	If Int(vCid)=0 Then Exit Function
	If Int(vCid)=0 Then TmpStr = ""
	SQL = "Select ClassParent From YSvoid_Class Where C_ID="&vCid
	Set Rs = YSvoid.Exec(SQL,1)
	If Not Rs.Eof Then
		ClassParent = Rs("ClassParent")
	End If
	Rs.Close

	If Int(Request("vCid")) = Int(ClassParent) Then
	For Ni=0 To ClassNum
		If Int(ClassArr(0,Ni))=Int(vCid) Then
			nCid=Int(ClassArr(6,Ni))
		End If
		If Int(ClassArr(0,Ni))<>Int(vCid) And Int(ClassArr(6,Ni))=nCid Then
			TmpStr=TmpStr & "," & ClassArr(0,Ni)
		End If
	Next
	Else
		TmpStr = vCid
	End If

	If Int(Request("vCid")) = Int(ClassParent) Then
			TmpStr="( ClassID in(" & vCid & TmpStr & ") )"
	Else
			TmpStr="( ClassID in(" & TmpStr & ") )"
	End If
	Class_Cid_Get2=TmpStr
End Function

'//* 分类列表(左) *//
Function Class_Left_Type(lt_url,lt_type,Depths)
  dim temp1,tmpimg,si,ni
  dim nCid, nPid, nDid, nIid, nTid, nHid
  if not ClassTrue then
    temp1="暂无分类"
    Class_Left_Type=temp1
    exit function
  End If
  if int(cid)=0 then    'if int(cid)=0 Or Depths=1 then
    for si=0 to ClassNum
      if int(ClassArr(2,si))=0 then
        nCid=int(ClassArr(0,si))
        tmpimg="sort_c1"
        temp1=temp1&vbcrlf&"<tr><td>"&YSvoid.img(tmpimg)&"&nbsp;&nbsp;<a href='"&lt_url&"c_id="&nCid&"'>"&ClassArr(1,si)&"</a></td></tr>"
        nCid=int(ClassArr(0,si))
      End If
      if int(ClassArr(2,si))=nCid And lt_type=1 And (Depths=0 Or Depths>1) then
        tmpimg="sort_c1"
	    temp1=temp1&vbcrlf&"<tr><td>&nbsp;&nbsp;&nbsp;&nbsp;"&YSvoid.img(tmpimg)&"&nbsp;&nbsp;<a href='"&lt_url&"c_id="&ClassArr(0,si)&"'>"&ClassArr(1,si)&"</a></td></tr>"
      End If
    next
    Class_Left_Type=temp1
    exit function
  End If
  for si=0 to ClassNum
    nCid=999999999
    nPid=999999999
    nDid=999999999
    nIid=0
    nTid=0
    nHid=0
	if int(ClassArr(0,si))=int(cid) then
		nCid=int(ClassArr(0,si))
		nPid=int(ClassArr(2,si))
		nDid=int(ClassArr(4,si))-1
		nIid=int(ClassArr(5,si))
		If Instr(ClassArr(3,si),",")>0 Then
			nTid=Int(Split(ClassArr(3,si),",")(nDid))
			nHid=nDid-1
			If nHid<0 Then nHid=0
			nHid=Int(Split(ClassArr(3,si),",")(nHid))
		End If
		exit for
	End If
  next
  for si=0 to ClassNum
    if nIid=0 then
      if int(ClassArr(4,si))=nDid then
        if int(ClassArr(0,si))=nTid or int(ClassArr(2,si))=nHid or int(ClassArr(2,si))=0 then
          tmpimg="sort_c1"
          if int(cid)=int(ClassArr(0,si)) then
            tmpimg="sort_c2"
          End If
          temp1=temp1&vbcrlf&"<tr><td>"&YSvoid.img(tmpimg)&"&nbsp;&nbsp;<a href='"&lt_url&"c_id="&ClassArr(0,si)&"'>"&ClassArr(1,si)&"</a></td></tr>"
        End If
      End If
      if int(ClassArr(2,si))=nPid then
        tmpimg="sort_c1"
        if int(cid)=int(ClassArr(0,si)) then
          tmpimg="sort_c2"
        End If
        if nPid=0 then
          temp1=temp1&vbcrlf&"<tr><td>"&YSvoid.img(tmpimg)&"&nbsp;&nbsp;<a href='"&lt_url&"c_id="&ClassArr(0,si)&"'>"&ClassArr(1,si)&"</a></td></tr>"
        else
          temp1=temp1&vbcrlf&"<tr><td>&nbsp;&nbsp;&nbsp;&nbsp;"&YSvoid.img(tmpimg)&"&nbsp;&nbsp;<a href='"&lt_url&"c_id="&ClassArr(0,si)&"'>"&ClassArr(1,si)&"</a></td></tr>"
        End If
      End If
    else
     If int(ClassArr(4,si))<Depths Or Depths=0 Then
      if int(ClassArr(2,si))=nPid then
        tmpimg="sort_c1"
        if int(cid)=int(ClassArr(0,si)) then
          tmpimg="sort_c2"
        End If
        temp1=temp1&vbcrlf&"<tr><td>"&YSvoid.img(tmpimg)&"&nbsp;&nbsp;<a href='"&lt_url&"c_id="&ClassArr(0,si)&"'>"&ClassArr(1,si)&"</a></td></tr>"
      End If
      if int(ClassArr(2,si))=int(cid) then
        tmpimg="sort_c1"
        if int(cid)=int(ClassArr(0,si)) then
          tmpimg="sort_c2"
        End If
        temp1=temp1&vbcrlf&"<tr><td>&nbsp;&nbsp;&nbsp;&nbsp;"&YSvoid.img(tmpimg)&"&nbsp;&nbsp;<a href='"&lt_url&"c_id="&ClassArr(0,si)&"'>"&ClassArr(1,si)&"</a></td></tr>"
      End If
     End If
    End If
 next
  Class_Left_Type=temp1
End Function

'//* 分类列表(下拉菜单) *//
Function Class_Select_Type(t_tit,t1,t2)
  dim si,temp1,ncid
  dim classid,nfid,classn
  temp1="<select name=c_id size=1>" & _
        "<option value=''>请选择"&t_tit&"类型</option>"
  if not ClassTrue then
    temp1=temp1&"</select>"
    Class_Select_Type=temp1
    exit function
  End If
  ncid=0
  for si=0 to ClassNum
    nfid=ClassArr(0,si)
    if int(ClassArr(2,si))=0 then
      temp1=temp1&vbcrlf&"<option class=bg_2 value='"&nfid&"'"
      if int(nfid)=int(cid) then temp1=temp1&" selected"
      temp1=temp1&">"&ClassArr(1,si)&"</option>"
      nfid=ClassArr(0,si)
    else
		If ClassArr(4,si)<t1 Or t1=0 Then
			temp1=temp1&vbcrlf&"<option value='"&nfid&"'"
			if int(nfid)=int(cid) then temp1=temp1&" selected"
			temp1=temp1&">"
			for i=1 to ClassArr(4,si)-1
				temp1=temp1&"│"
			next
			temp1=temp1&"├"&ClassArr(1,si)
			temp1=temp1&"</option>"
		End If
	End If
  next
  temp1=temp1&"</select>"
  Class_Select_Type=temp1
End Function

'//* 分类结束 *//
Sub Class_End()
	If IsArray(ClassArr) Then
		Erase ClassArr
	End If
End Sub

'//* Cid初始 *//
Sub Cid_Load()
	Cid=YSvoid.Code_ID("c_id")
End Sub

function sel_mod_num(n1,n2)
  sel_mod_num=false
  dim nn1
  nn1=n2/n1
  if instr(nn1,".")=0 then
    sel_mod_num=true
  End If
End Function

'//* SQL拼接 *//
Sub Cid_SQL(csst,csstt)
  dim temp
  temp=csstt
  if cid>0 then
    sqladd=sqladd&" and c_id="&cid
    pageurl=pageurl&"c_id="&cid&"&"
  End If
  if csst=1 or csst=2 then
    if len(keyword)>0 then
      sqladd=sqladd&" and "&temp&" like '%"&keyword&"%'"
      pageurl=pageurl&"keyword="&server.urlencode(keyword)&"&"
      if csst=2 then pageurl=pageurl&"sea_type="&sea_type&"&"
    End If
  End If
End Sub

'//* 当前标题所在分类 *//
Function Class_Name_Topic(vCid)
  if Int(vCid)=0 Then Exit Function
  Class_Name_Topic="<font class=gray>[<a href=List.asp?c_id="&vCid&">"&Class_Name_Get(vCid)&"</a>]</font>"
End Function

%>