<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="../user/checklogin.asp"-->
<%
action=request("action")
Content=trim(request.form("Content"))
id=request("id")
quid=request("quid")

if action="add" then 
       
sql = "select * from "&jieducm&"_register where username='"&session("useridname")&"' and tribeid="&id&""
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if rs.eof then
	Response.Write("<script>alert('��û�м���˲��䲻�����ԣ�');window.location.href='tribeinfo.asp?id="&id&"';</script>")
    response.End()			
end if

if content="" then
	Response.Write("<script>alert('�������ݲ���Ϊ�գ�');window.location.href='tribeinfo.asp?id="&id&"';</script>")
    response.End()	
end if
   
	Sql= "select * from "&jieducm&"_tribebook"
	Set rs=server.createobject("ADODB.RECORDSET")
	Rs.Open Sql,conn,3,3			
	rs.addnew
	rs("username")=session("useridname")
    rs("tribeid")=id
	rs("content")=content
	rs("now")=now()
    rs.update
    rs.close	
    set rs=nothing
	 Response.Write("<script>alert('���Գɹ�!');window.location.href='tribeinfo.asp?id="&id&"';</script>")
	 response.End()
	 
elseif action="del" then

sql = "select * from "&jieducm&"_tribe where id="&quid&""
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
 if session("useridname")<>rs("username") then
	Response.Write("<script>alert('��û��Ȩ��ɾ�������ԣ�');window.location.href='tribeinfo.asp?id="&quid&"';</script>")
    response.End()	
 end if		
end if

sql="delete from "&jieducm&"_tribebook where id in("&id&")"
conn.execute(sql)


Response.Write("<script>alert('ɾ���ɹ�!');window.location.href='tribeinfo.asp?id="&quid&"';</script>")
response.End()
end if

%>