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
	Response.Write("<script>alert('你没有加入此部落不能留言！');window.location.href='tribeinfo.asp?id="&id&"';</script>")
    response.End()			
end if

if content="" then
	Response.Write("<script>alert('留言内容不能为空！');window.location.href='tribeinfo.asp?id="&id&"';</script>")
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
	 Response.Write("<script>alert('留言成功!');window.location.href='tribeinfo.asp?id="&id&"';</script>")
	 response.End()
	 
elseif action="del" then

sql = "select * from "&jieducm&"_tribe where id="&quid&""
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
 if session("useridname")<>rs("username") then
	Response.Write("<script>alert('你没有权限删除此留言！');window.location.href='tribeinfo.asp?id="&quid&"';</script>")
    response.End()	
 end if		
end if

sql="delete from "&jieducm&"_tribebook where id in("&id&")"
conn.execute(sql)


Response.Write("<script>alert('删除成功!');window.location.href='tribeinfo.asp?id="&quid&"';</script>")
response.End()
end if

%>