
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="../md5.asp"-->

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
'Copyright (C) 2008��2011 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------

'�˹���С��ʹ�ã���ɾ�����ݿ���������Ϣ������Դ��������֮�����ʹ�ô˹���
if request("action")="ok" then
czm=request("czm")
if md5(czm)<>admin_czmpass then

  Set rs=server.createobject("ADODB.RECORDSET")
   rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
   rs.addnew
   rs("username")=session("AdminName")
   rs("class")="ϵͳ��ʼ��"
   rs("content")=session("AdminName")&"����Ա����ϵͳ��ʼ���������������"
   rs("jiegou")="ʧ��"
   rs.update
   rs.close
   Response.Write("<script>alert('����������!�벻Ҫ�ظ�������ƽ̨���¼�����Ϊ��');history.back();</script>")
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
Response.Write("<script>alert('ϵͳ��ʼ���ɹ�!');window.location.href='admin_index.asp';</script>")
end if
end if
%>
<form id="form1" name="form1" method="post" action="?action=ok">
  <label>
  �����룺<input type="password" name="czm" />
  </label>
  <label>
  <input type="submit" name="Submit" value="�ύ" />
  </label>
</form>