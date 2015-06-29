<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="checklogin.asp"-->
<%
dim theid,thdel
id=request("adid")

if id="" then
 Response.Write("<script>alert('请先选择要删除的记录!');window.location.href='sms.asp';</script>")
 response.End()
end if
sql="delete from "&jieducm&"_sms where id in("&id&") and username='"&session("useridname")&"'"
conn.execute(sql) 

conn.close
set conn=nothing
 Response.Write("<script>alert('操作成功!');window.location.href='sms.asp';</script>")
 response.End()
%>