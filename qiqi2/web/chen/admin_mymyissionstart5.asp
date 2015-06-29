
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
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
'Copyright (C) 2008－2009 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
id=request("id")
action=request.Form("action")
num=request.Form("num")
idd=request.Form("idd")
if action="ok" then
if num="" then
 Response.Write("<script>alert('设置已完成收藏个数不能为空!');history.back();</script>")
response.End()
end if
sqlr="update "&jieducm&"_zhongxin set jieshou_num="&num&"  where id="&idd&" "
conn.execute sqlr

Response.Write("<script>alert('设置成功!');window.location.href='admin_MyMission5.asp';</script>")
response.End()
end if 

%>
<form id="form1" name="form1" method="post" action="?">
<input name="idd" type="hidden" value="<%=id%>" />
<input name="action" type="hidden" value="ok" />
设置已完成收藏个数：
  <label>
  <input type="text" name="num"  onKeyUp="if(isNaN(value))execCommand('undo')" /> 只能是数字
  </label>
  <label>
  <input type="submit" name="Submit" value="提交" />
  </label>
</form>