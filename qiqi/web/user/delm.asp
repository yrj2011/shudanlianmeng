<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="checklogin.asp"-->
<%

dim theid,thdel
theid=request("adid")
thedel=request("del")

if theid="" or thedel="" then
response.write "<script>alert('参数错误，关闭窗口！');window.close();</script>"
response.end
end if


if thedel="data" then
sql="delete from "&jieducm&"_china_data where adid in("&theid&")"
conn.execute(sql) 

sql="delete from "&jieducm&"_china_message where adid in("&theid&")"
conn.execute(sql)
conn.execute(sql)
'''''''''''删除html开始''''''''''''''''
Set fdel = CreateObject("Scripting.FileSystemObject")	
for i=1 to request("adid").count
id=request("adid")(i) 
              tempImageDir = Server.MapPath("html/"&id&".htm")
	             fdel.DeleteFile(tempImageDir)
        next
       set fdel=nothing
'''''''''''删除html结束'''''''''''''''
end if

if thedel="message" then
set rs=server.createobject("adodb.recordset")
rs.open "select * from "&jieducm&"_china_message where id in("&theid&") order by adid desc",conn,1,1
dim adid()
a=0
for i=0 to rs.recordcount-1
redim Preserve adid(a)
if ad=rs("adid") then
rs.movenext
else
adid(a)=rs("adid")
ad=rs("adid")
a=a+1
rs.movenext
end if
next
rs.close
set rs=nothing
sql="delete from "&jieducm&"_china_message where id in("&theid&")"
conn.execute(sql)
conn.close
set conn=nothing

end if

                             
response.write"<script>alert('删除成功！');location='"&request.ServerVariables("HTTP_REFERER")&"'</script>"
call closeconn()
%>