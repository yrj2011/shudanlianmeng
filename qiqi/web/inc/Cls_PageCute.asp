<%

Class Cls_PageCute
	Public pageurl							'��ǰ����
	Public pagetarget                       '��ǰ����ģʽ
	Public nummer							'ÿҳ��ʾ����
	Public page_t							'��ʾ���ͣ�����ĸ�ʽ��ʽ
	Public page_p							'��ҳ����,viewpage
	Public page_unit						'��λע�ͣ�����ƪ��
	Public page_tit							'������ƣ����¡����š��û���
	Public page_color						'����ɫ
	Public rssum							'ȡ���ݿ�������ֵ
	Public viewpage,thepages				'��ǰҳ
	Public page_num							'���ݿ�ȡ�������ݸ���
	Public page_now							'��ǰ��ʾ�ķ�ҳ����
	Public v_id								'����ID����
	Public sql_count						'ȡ���ݿ���������SQL���ƴ��
	Public sql_master						'��ƴ�ӹ���SQL����ҳ���
	Private mList,mTable,mAdd,mOrder

	Private Sub Class_Initialize()
		pageurl="?"
		pagetarget=""
		nummer=20
		page_t=1
		page_p=5
		page_unit="��"
		page_tit=tit_fir
		page_color="#ff0000"
		v_id="id"
		sql_count=""
		sql_master=""
		rssum=0
		mList=""
		mTable=""
		mAdd=""
		mOrder=""
	End Sub

'//* ��ʼ�� Page *//
	Private Property Get Page
		 Page = Trim(Request("page"))
		 'Page = Trim(Request.QueryString("page"))
		 If Not(IsNumeric(Page)) Then Page = 1
	End Property
	'//* ��һҳ��ť��ʾ��ʽ *//
	Private Property Get Btn_First
		Btn_First = "<font face=""webdings"">9</font>"
	End Property
	'//* ǰһҳ��ť��ʾ��ʽ *//
	Private Property Get Btn_Prev
		Btn_Prev = "<font face=""webdings"">7</font>"
	End Property
	'//* ��һҳ��ť��ʾ��ʽ *//
	Private Property Get Btn_Next
		Btn_Next = "<font face=""webdings"">8</font>"
	End Property
	'//* ���һҳ��ť��ʾ��ʽ *//
	Private Property Get Btn_Last
		Btn_Last = "<font face=""webdings"">:</font>"
	End Property

		'--- SQL����ʼ��
	Public Sub p_init(vList,vTable,vAdd,vOrder)
		If vAdd<>"" Then
			sql_count="select count("&v_id&") from "&vTable&" where ("&vAdd&")"
		else
			sql_count="select count("&v_id&") from "&vTable
		End If

		mList=vList
		mTable=vTable
		mAdd=vAdd
		mOrder=vOrder
	End Sub

	Public Sub p_info()
		If rssum Mod nummer > 0 Then
			thepages=(rssum\nummer)+1
		Else
			thepages=rssum\nummer
		End If
		If int(page)>int(thepages) Or int(page)<1 Then
			viewpage=1
		Else
			viewpage=int(page)
		End If
		if page_t=3 then exit sub
		page_now=Int((viewpage-1)*nummer)
		'page_now=viewpage*nummer
		page_num=nummer
		if int(viewpage*nummer)>int(rssum) then
			page_num=Int(nummer-(viewpage*nummer-rssum))
			'page_now=page_num
		End If
	End Sub

	'--- SQL���ƴ�ӣ���ҳȡ���ݵ�SQL
	Public Sub p_main(vList,vTable,vAdd,vOrder)
	Call p_info()
		If vAdd<>"" Then
			sql="select top "&nummer*viewpage&" "&VList&" from "&VTable&" where ("&VAdd&") order by "&VOrder
			Call YSvoid.Exe_Conn(Rs,SQL,3)
        else
            sql="select top "&nummer*viewpage&" "&VList&" from "&VTable&" order by "&VOrder
			Call YSvoid.Exe_Conn(Rs,SQL,3)
		End If
			if int(viewpage)>1 then rs.move (viewpage-1)*nummer
if int(viewpage*nummer)>int(rssum) then page_num=nummer-(viewpage*nummer-rssum)
	End Sub

'����ַ���
	Public Property Get page_var
		page_var="����<font color="""&page_color&""">"&rssum&"</font>"&page_unit&"<font title='ÿҳ"&nummer&page_unit&page_tit&"'>"&page_tit&"</font>&nbsp;&nbsp;&nbsp;&nbsp;ҳ�Σ�<font color="""&page_color&""">"&viewpage&"</font>/<font color="""&page_color&""">"&thepages&"</font>"
	End Property

	'��ҳ���
	Public Property Get page_cute
		Select Case page_t
		Case 1
			page_cute=YSvoid_pagecute()
		Case Else
			page_cute=pagecute_fun()
		End Select
	End Property

    '���ٲ��ҷ�ҳ
	Public Property Get page_shortcut
	    page_shortcut="<td><input type=text class=txt name=page value='"&viewpage&"' size=3 style=""text-align:center;""></td><td><input type=button class=txt value='GO' onclick=""javascript:location.href='"&pageurl&"page='+page.value""></td>"
	End Property


Private Property Get YSvoid_pagecute()
	    if pagetarget<>"" then pagetarget="target="&pagetarget
		if int(thepages)=0 then
		  YSvoid_pagecute="<font color=" & page_color & "> "&btn_first&" <b>1</b> "&btn_last&" </font>"
		  exit Property
		End If
		dim pi,ppp,pl,pr
		pi=1
		ppp=page_p\2
		if page_p mod 2 = 0 then ppp=ppp-1
		pl=viewpage-ppp
		pr=pl+page_p-1
		if pl<1 then
		  pr=pr-pl+1
		  pl=1
		  if pr>thepages then pr=thepages
		End If
		if pr>int(thepages) then
		  pl=pl+thepages-pr
		  pr=thepages
		  if pl<1 then pl=1
		End If
		if int(viewpage)=1 then
		   YSvoid_pagecute=YSvoid_pagecute&" <font color=" & page_color & "> "&btn_first&" </font> "
		else
		   YSvoid_pagecute=YSvoid_pagecute&" <a href='"& pageurl &"page=1' title='��һҳ' "& pagetarget&">"&btn_first&"</a> "
		End If
		if int(viewpage)>1 then
		  YSvoid_pagecute=YSvoid_pagecute&" <a href='"& pageurl &"page="&viewpage-1&"' title='��һҳ' "& pagetarget&">"&btn_prev&"</a> "
		End If
		for pi=pl to pr
		  if cint(viewpage)=cint(pi) then
		    YSvoid_pagecute=YSvoid_pagecute&" <font color=" & page_color & " face=����><b>" & pi & "</b></font> "
		  else
		    YSvoid_pagecute=YSvoid_pagecute&" <a href='"& pageurl &"page="& pi &"' title='�� " & pi & " ҳ' clases=0 style=""font-family:����"" "& pagetarget&"><b>" & pi & "</b></a> "
	  End If
		next
		if cint(viewpage)<cint(thepages) then
		  YSvoid_pagecute=YSvoid_pagecute&" <a href='"& pageurl &"page="&viewpage+1&"' clases=0 title='��һҳ' "& pagetarget&">"&btn_next&"</a> "
		End If
		if cint(viewpage)=cint(thepages) then
		  YSvoid_pagecute=YSvoid_pagecute&" <font color=" & page_color & "> "&btn_last&" </font> "
		else
		  YSvoid_pagecute=YSvoid_pagecute&" <a href='"& pageurl &"page="& thepages &"' title='���һҳ' "& pagetarget&">"&btn_last&"</a> "
		End If
	end Property

	Private Property Get pagecute_fun()
	  dim pf1,pf2,pf3,pf4
	  pf1="��ҳ"    '"��һҳ"
	  pf2="��һҳ"
	  pf3="��һҳ"
	  pf4="βҳ"  '"���һҳ"
	  if pagetarget<>"" then pagetarget="target="&pagetarget
	  if cint(viewpage)=1 then
	    pagecute_fun="<font class=gray>"&pf1&"</font>&nbsp;<font class=gray>"&pf2&"</font>&nbsp;"
	  else
	    pagecute_fun=pagecute_fun&"<a href='"&pageurl&"page=1' title='"&pf1&"' "& pagetarget&">"&pf1&"</a>&nbsp;<a href='"&pageurl&"page="&cint(viewpage)-1&"' title='"&pf2&"' "& pagetarget&">"&pf2&"</a>&nbsp;"
	  End If
	  if cint(viewpage)=cint(thepages) or cint(thepages)=0 then
	    pagecute_fun=pagecute_fun&"<font class=gray>"&pf3&"</font>&nbsp;<font class=gray>"&pf4&"</font>&nbsp;"
	  else
	    pagecute_fun=pagecute_fun&"<a href='"&pageurl&"page="&cint(viewpage)+1&"' title='"&pf3&"' clases=0 "& pagetarget&">"&pf3&"</a>&nbsp;<a href='"&pageurl&"page="&cint(thepages)&"' title='"&pf4&"' "& pagetarget&">"&pf4&"</a>"
	  End If
	end Property
End Class


Public Function PageCute_Topic(ByVal sPageCuteNum,ByVal sReplyNum,ByVal sViewUrl)
  Dim PageCutePage,PageCutei,PageCuteColor,temp1
  PageCutePage=0
  PageCutei=0
  PageCuteColor="red2"
  PageCutePage=PageCute_Num(sPageCuteNum,sReplyNum)
  temp1=""
  If (PageCutePage*sPageCuteNum)<sReplyNum Then PageCutePage=PageCutePage+1
  If PageCutePage>1 Then
	temp1=temp1&vbcrlf&"&nbsp;<img src='"&YSvoid.web_dir_skin&"forum/page_head.gif' align=absMiddle title='���ٷ�ҳ' border=0>&nbsp;"
    For PageCutei=2 To 4
      If PageCutei>PageCutePage Then
	    Exit For
	  End If
      temp1=temp1&vbcrlf&" <a href='"&sViewUrl&"&page="&PageCutei&"'><font class="&PageCuteColor&" title='�� "&PageCutei&" ҳ' style=""font-family:����;font-weight: bold"">"&PageCutei&"</font></a> "
    Next
    If PageCutePage>4 Then
      temp1=temp1&vbcrlf&" <font class="&PageCuteColor&">�� </font> <a href='"&sViewUrl&"&page="&PageCutePage&"'><font class="&PageCuteColor&" title='�� "&PageCutePage&" ҳ' style=""font-family:����;font-weight: bold"">"&PageCutePage&"</font></a> "
	End If
  End If
  PageCute_Topic=temp1
End Function

Public Function PageCute_Num(ByVal sPageNum,ByVal sView)
  PageCute_Num=sView/sPageNum
  PageCute_Num=int(PageCute_Num)+1
End Function

%>