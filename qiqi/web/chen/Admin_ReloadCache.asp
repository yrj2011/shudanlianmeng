<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Response.Buffer = True
%>
<!--#Include File="inc/function.asp"-->
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：捷度传媒
'官方主页：http://www.jieducm.cn
'技术支持：robin 
'QQ;616591415
'电话：13598006137
'版权：版权属于捷度传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2008－2011 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
Dim CacheName,nt2003
Set nt2003= New System_Cls
CacheName=nt2003.CacheName
Call delallcache()
'session("AdminName")=""
'Response.Cookies("asp163")("UserName")=""
'Response.Cookies("asp163")("UserLevel")=""
 Response.Redirect "index.asp"

Sub delallcache()
	Dim cachelist,i
	Cachelist=split(GetallCache(),",")
	If UBound(cachelist)>1 Then
		For i=0 to UBound(cachelist)-1
			DelCahe Cachelist(i)
			Response.Write "更新 <b>"&Replace(cachelist(i),CacheName&"_","")&"</b> 完成<br>"		
		Next
		Response.Write "更新了"
		Response.Write UBound(cachelist)-1
		Response.Write "个缓存对象<br>"	
	Else
		Response.Write "所有对象已经更新。"
	End If
End Sub 
Function  GetallCache()
	Dim Cacheobj
	For Each Cacheobj in Application.Contents
	If CStr(Left(Cacheobj,Len(CacheName)+1))=CStr(CacheName&"_") Then	
		GetallCache=GetallCache&Cacheobj&","
	End If
	Next
End Function
Sub DelCahe(MyCaheName)
	Application.Lock
	Application.Contents.Remove(MyCaheName)
	Application.unLock
End Sub
%>