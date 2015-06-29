<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：捷度传媒
'官方主页：http://www.jieducm.cn
'技术支持：robin 
'QQ;616591415
'电话：13598006137
'版权：版权属于捷度传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2008－2011 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------

dim theid,thdel
theid=request("adid")
thedel=request("del")

	   Set rsr=server.createobject("ADODB.RECORDSET")
	   rsr.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rsr.addnew
		rsr("username")=session("AdminName")
		rsr("class")="删除站内信息"
		rsr("content")="管理员删除了站内信息"
		rsr("jiegou")="成功"
    	rsr.update
    	rsr.close
if theid="" or thedel="" then
response.write "<script>alert('参数错误，关闭窗口！');window.close();</script>"
response.end
end if

if thedel="user" then
sql="delete from "&jieducm&"_china_user where instr(1,'"&theid&"',uid)>0"
conn.execute(sql) 

sql="delete from "&jieducm&"_china_data where instr(1,'"&theid&"',uid)>0"
conn.execute(sql)

sql="delete from "&jieducm&"_china_message where instr(1,'"&theid&"',uid)>0"
conn.execute(sql)

'''''''''''删除html开始''''''''''''''''

end if

if thedel="data" then
sql="delete from "&jieducm&"_china_data where adid in("&theid&")"
conn.execute(sql) 

sql="delete from "&jieducm&"_china_message where adid in("&theid&")"
conn.execute(sql)

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

if thedel="pass" then
conn.execute("update "&jieducm&"_china_data set mark='yes' where adid in("&theid&")") 
end if
                             
response.redirect request.ServerVariables("HTTP_REFERER")
%>