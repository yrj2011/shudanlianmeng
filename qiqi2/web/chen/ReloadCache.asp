<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Response.Buffer = True
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<%
Dim CacheName,nt2003
Set nt2003= New System_Cls
CacheName=nt2003.CacheName
%>
<html>
<head>
<title>�������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Admin_Style.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" Class="border" >
  <tr>       
    <td height="22" align="center" colspan="3" class="title" ><b>�� �� �� �� �� ��</b></td>
  </tr>
  <% Call delallcache()%>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" colspan="3" align="center" background="../images/admin_bg_bottom.gif"><font color=#cccccc>Copyright &copy; 2013</font> <a href='http://www.778892.com' target="_blank"></a> <font color="#cccccc">. All Rights Reserved.</font></td>
  </tr>
</table>
</body>
</html>

<%
Sub delallcache()
	Dim cachelist,i
	Cachelist=split(GetallCache(),",")
	If UBound(cachelist)>1 Then
	    Response.Write "<tr class='tdbg'><td width='8%'><b>���</b></td><td width='20%'><b>���¶���</b></td><td><b>״̬</b></td></tr>"
		For i=0 to UBound(cachelist)-1
			DelCahe Cachelist(i)
			Response.Write "<tr class='tdbg'><td align='left'>"
			Response.Write i+1 & "</td><td><font color='#FF6600'>"&Replace(cachelist(i),CacheName&"_","")&"</font></td><td>���</td></tr>"		
		Next
		Response.Write "<tr class='tdbg' align='center'><td colspan='3' align='center'><b>�����޸Ĺ�������&nbsp;&nbsp;<font color='#FF6600'>"
		Response.Write UBound(cachelist) 'ѩ��ȥ�� -1
		Response.Write "</font>&nbsp;&nbsp;���������</b></td></tr>"	
	Else
		Response.Write "<tr class='tdbg'><td colspan='3' align='center'>���ж����Ѿ����¡�</td></tr>"
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