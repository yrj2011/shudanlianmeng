<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<%
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../jieducm/#jieducm.asp")
conn.Open connstr
If Err Then
	err.Clear
	Response.Write "���ݿ����ӳ������������ִ���"
	Response.End
End If


Dim ip,url,sql
ip = request("ip")
      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From Rc_SqlIn where Sqlin_IP='"&ip&"' and Kill_ip" ,Conn,3,3  
	  if not rs.eof then
	 	 Response.Write("<script>alert('��IP�Ѿ�������!');history.back();</script>")
		 response.End()
	 end if

	 
	sql = "insert into Rc_SqlIn(Sqlin_IP,Kill_ip) values('"&ip&"',1)"
	'response.write sql
	conn.Execute(sql)
	conn.close
	Set conn = Nothing
	response.write "<script>alert('���óɹ���"&ip&" ��IP�ѱ�ϵͳ������');window.location.href='usermanage.asp';</script>"
	response.End()
		
%>
