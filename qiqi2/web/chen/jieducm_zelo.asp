
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="../md5.asp"-->

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

'此功能小心使用，将删除数据库里所有信息。购买源码后测试完之后可以使用此功能
if request("action")="ok" then
czm=request("czm")
if md5(czm)<>admin_czmpass then

  Set rs=server.createobject("ADODB.RECORDSET")
   rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
   rs.addnew
   rs("username")=session("AdminName")
   rs("class")="系统初始化"
   rs("content")=session("AdminName")&"管理员进行系统初始化操作码输入错误"
   rs("jiegou")="失败"
   rs.update
   rs.close
   Response.Write("<script>alert('操作码有误!请不要重复操作，平台会记录你的行为！');history.back();</script>")
   response.End()
else
conn.execute("delete from "&jieducm&"_zhongxin")
conn.execute("delete from "&jieducm&"_xinyu")
conn.execute("delete from "&jieducm&"_toushu")
conn.execute("delete from "&jieducm&"_tixian")
conn.execute("delete from "&jieducm&"_register")
conn.execute("delete from "&jieducm&"_recordm")
conn.execute("delete from "&jieducm&"_record")
conn.execute("delete from "&jieducm&"_qq")
conn.execute("delete from "&jieducm&"_jieshou")
conn.execute("delete from "&jieducm&"_hei")
conn.execute("delete from "&jieducm&"_guestbook")
conn.execute("delete from "&jieducm&"_friendlinks")
conn.execute("delete from "&jieducm&"_dj586_pay")
conn.execute("delete from "&jieducm&"_dj586_Jicar")
conn.execute("delete from "&jieducm&"_dj586_admin")
conn.execute("delete from "&jieducm&"_ding")
conn.execute("delete from "&jieducm&"_depot")
conn.execute("delete from "&jieducm&"_chongjilu")
conn.execute("delete from "&jieducm&"_china_message")
conn.execute("delete from "&jieducm&"_chengfa")
conn.execute("delete from "&jieducm&"_beifei")
conn.execute("delete from "&jieducm&"_keeper")
conn.execute("delete from "&jieducm&"_tribebook")
conn.execute("delete from "&jieducm&"_tribe")
conn.execute("delete from "&jieducm&"_sms")
conn.execute("delete from "&jieducm&"_pay")
Response.Write("<script>alert('系统初始化成功!');window.location.href='admin_index.asp';</script>")
end if
end if
%>
<form id="form1" name="form1" method="post" action="?action=ok">
  <label>
  操作码：<input type="password" name="czm" />
  </label>
  <label>
  <input type="submit" name="Submit" value="提交" />
  </label>
</form>