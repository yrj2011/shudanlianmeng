<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Response.Buffer = True
%>
<!--#Include File="inc/function.asp"-->
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
			Response.Write "���� <b>"&Replace(cachelist(i),CacheName&"_","")&"</b> ���<br>"		
		Next
		Response.Write "������"
		Response.Write UBound(cachelist)-1
		Response.Write "���������<br>"	
	Else
		Response.Write "���ж����Ѿ����¡�"
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