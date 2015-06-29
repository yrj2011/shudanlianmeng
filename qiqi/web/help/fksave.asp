<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
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
if session("useridname")="" then
Response.Write("<script>alert('您还没有登录不能留言,请先登录!');window.location.href='../login.asp';</script>")
response.End()
end if
if request.form("name")="" then
response.Write "<script LANGUAGE='javascript'>alert('请填写您的姓名');history.go(-1);</script>"
response.End
end if
if request.form("content")="" then
response.Write "<script LANGUAGE='javascript'>alert('请填写您的留言与建议');history.go(-1);</script>"
response.End
end if

if request.form("temp")="" then
response.write "<script language='javascript'>" & VbCRlf
response.write "alert('非法操作!');" & VbCrlf
response.write "history.go(-1);" & vbCrlf
response.write "</script>" & VbCRLF
'由于表单被重复提交(标志为session("antry")为空)引起
else
iname=HtmlEncode(request.Form("name"))
iqq=HtmlEncode(request.Form("qq"))
iemail=HtmlEncode(request.Form("email"))
iurl=HtmlEncode(request.form("url"))
isex="boy"
icontent=HtmlEncode(request.form("content"))

set rs=server.createobject("adodb.recordset")
rs.open "select * from "&jieducm&"_guestbook where (id is null)",conn,1,3
rs.addnew
rs("name")=iname
rs("sex")=isex
rs("content")=icontent
rs("time")=now()
rs.update
rs.close
set rs=nothing
session("antry")="" '提交成功，清空session("antry")，以防重复提交！！
end if
%>
<meta http-equiv="refresh" content="1;URL=viewreturn.asp">
<title>留言成功</title>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<table width="62%" border="0" align="center">
  <tr>
    <td align="center"><p> <img src="images/loading.gif" width="94" height="27"><br>
        <br>
        <font size="2">留言成功，1秒钟后返回！</font></p>
      </td>
  </tr>
</table>