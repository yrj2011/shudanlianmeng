<%
'Response.write(Session("AdminName"))
If Session("AdminName")="" Then
	Response.Write("�Բ�������Ȩ�޲���")
	Response.End()
End If
%>