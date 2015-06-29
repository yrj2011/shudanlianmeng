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
	Response.Write "数据库连接出错，请检查连接字串。"
	Response.End
End If


Dim ip,url,sql
ip = request("ip")
      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From Rc_SqlIn where Sqlin_IP='"&ip&"' and Kill_ip" ,Conn,3,3  
	  if not rs.eof then
	 	 Response.Write("<script>alert('此IP已经被锁定!');history.back();</script>")
		 response.End()
	 end if

	 
	sql = "insert into Rc_SqlIn(Sqlin_IP,Kill_ip) values('"&ip&"',1)"
	'response.write sql
	conn.Execute(sql)
	conn.close
	Set conn = Nothing
	response.write "<script>alert('设置成功！"&ip&" 此IP已被系统锁定！');window.location.href='usermanage.asp';</script>"
	response.End()
		
%>
