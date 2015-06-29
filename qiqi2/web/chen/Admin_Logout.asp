<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Response.Buffer = True
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#Include File="inc/function.asp"-->
<%

Set nt2003= New System_Cls
CacheName=nt2003.CacheName
Call delallcache()
session("AdminName")=""
Session.Contents.RemoveAll()
Response.Cookies("asp163")("UserName")=""
Response.Cookies("asp163")("UserLevel")=""
Response.Redirect "Admin_Login.asp"
'Application.Contents.RemoveAll()
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