<%
'Response.write(Session("AdminName"))
If Session("AdminName")="" Then
	Response.Write("对不起，您的权限不够")
	Response.End()
End If
%>