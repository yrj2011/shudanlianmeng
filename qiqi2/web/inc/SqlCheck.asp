<%
Fy_In = "'��and��exec��insert��select��delete��update��count��chr��mid��master��truncate��char��declare��or"
Fy_Inf = split(Fy_In,"��")
If Request.Form<>"" Then
For Each Fy_Post In Request.Form
For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh))<>0 Then
Response.Write "<Script Language=JavaScript>alert('�����ʵ�ҳ�治���ڣ�');</Script>"
Response.End()
End If
Next
Next
End If
If Request.QueryString<>"" Then
For Each Fy_Get In Request.QueryString
For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh))<>0 Then
Response.Write "<Script Language=JavaScript>alert('�����ʵ�ҳ�治���ڣ�');</Script>"
Response.End()
End If
Next
Next
End If
 
%> 