<!--#include file="inc/index_conn.asp"-->
<%
	UserName = Request.QueryString("n")
	Sql = "select * from "&jieducm&"_register where username='"&UserName&"'"
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	IF Rs.Eof or Rs.Bof Then
		Response.Write("document.all["&"""c0"""&"].innerHTML="&"""<font color=#ff0000>�̹�ϲ�����û�������ע��</font>"""&";")		
		'Response.Write("document.all["&"""c0"""&"].innerHTML="&""""&UserName&""""&";")
	Else
		'Response.Write("document.all["&"""c0"""&"].innerHTML="&"""����"""&";")
		Response.Write("document.all["&"""c0"""&"].innerHTML="&"""<font color=#ff0000>���Բ����û���"&UserName&"�Ѿ�����</font>&nbsp;&nbsp;"""&";")
	End IF

%>
