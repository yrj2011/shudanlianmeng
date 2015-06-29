<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->


<%
id=request("id")
if id="" then
 Response.Write("<script>alert('请先选择要删除的记录!');history.back();</script>")
 response.End()
end if
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_qq where id in ("&id&")",conn,3,3
set rs=nothing
conn.close
set conn=nothing
  response.Redirect("qqadmin.asp")
  %>