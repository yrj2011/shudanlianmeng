
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷����ݶȴ�ý
'�ٷ���ҳ��http://www.jieducm.cn
'����֧�֣�robin 
'QQ;616591415
'�绰��13598006137
'��Ȩ����Ȩ���ڽݶȴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2008��2009 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
id=request("id")
action=request.Form("action")
num=request.Form("num")
idd=request.Form("idd")
if action="ok" then
if num="" then
 Response.Write("<script>alert('����������ղظ�������Ϊ��!');history.back();</script>")
response.End()
end if
sqlr="update "&jieducm&"_zhongxin set jieshou_num="&num&"  where id="&idd&" "
conn.execute sqlr

Response.Write("<script>alert('���óɹ�!');window.location.href='admin_MyMission5.asp';</script>")
response.End()
end if 

%>
<form id="form1" name="form1" method="post" action="?">
<input name="idd" type="hidden" value="<%=id%>" />
<input name="action" type="hidden" value="ok" />
����������ղظ�����
  <label>
  <input type="text" name="num"  onKeyUp="if(isNaN(value))execCommand('undo')" /> ֻ��������
  </label>
  <label>
  <input type="submit" name="Submit" value="�ύ" />
  </label>
</form>