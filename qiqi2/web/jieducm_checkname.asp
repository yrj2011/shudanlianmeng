<!--#include file="inc/index_conn.asp"-->
<%
	UserName = Request.QueryString("n")
	Sql = "select * from "&jieducm&"_register where username='"&UserName&"'"
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	IF Rs.Eof or Rs.Bof Then
		Response.Write("document.all["&"""c0"""&"].innerHTML="&"""<font color=#ff0000>√恭喜您此用户名可以注册</font>"""&";")		
		'Response.Write("document.all["&"""c0"""&"].innerHTML="&""""&UserName&""""&";")
	Else
		'Response.Write("document.all["&"""c0"""&"].innerHTML="&"""存在"""&";")
		Response.Write("document.all["&"""c0"""&"].innerHTML="&"""<font color=#ff0000>×对不起用户名"&UserName&"已经存在</font>&nbsp;&nbsp;"""&";")
	End IF

%>
