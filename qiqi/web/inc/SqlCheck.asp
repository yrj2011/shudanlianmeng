<%
Fy_In = "'飞and飞exec飞insert飞select飞delete飞update飞count飞chr飞mid飞master飞truncate飞char飞declare飞or"
Fy_Inf = split(Fy_In,"飞")
If Request.Form<>"" Then
For Each Fy_Post In Request.Form
For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh))<>0 Then
Response.Write "<Script Language=JavaScript>alert('您访问的页面不存在！');</Script>"
Response.End()
End If
Next
Next
End If
If Request.QueryString<>"" Then
For Each Fy_Get In Request.QueryString
For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh))<>0 Then
Response.Write "<Script Language=JavaScript>alert('您访问的页面不存在！');</Script>"
Response.End()
End If
Next
Next
End If
 
%> 