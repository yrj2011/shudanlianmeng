
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<%
if request.Form("action")="ok" then
tdate=request.Form("startdate")
mdate=request.Form("enddate")
 if tdate="" or mdate="" then
 	 Response.Write("<script>alert('��ʼ�ͽ������ڶ�����Ϊ��!');history.back();</script>")
     response.End()
 end if
 
 if tdate="��ѡ������" or mdate="��ѡ������" then
  	 Response.Write("<script>alert('��ʼ�ͽ������ڶ�����Ϊ��!');history.back();</script>")
     response.End()
 end if
 
sql="select count(*) as yu  from "&jieducm&"_record where now between '"&tdate&"' and '"&mdate&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF not Rs.Eof Then
countnum=rs("yu")
else
  	 Response.Write("<script>alert('���ݿ����޴����ڷ�Χ�ڵ�����!');history.back();</script>")
     response.End()
end if

sqlfy="delete  from "&jieducm&"_record where now between '"&tdate&"' and '"&mdate&"'"
Conn.execute(sqlfy)

Response.Write("<script>alert('���ι�ɾ��"&countnum&"����¼!');window.location.href='admin_delrecord.asp';</script>")
response.End()
	 
end if
%>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
<script src="inc/jsdate.js" type="text/javascript"></script>

<form name="form1" method="post" action="admin_delrecord.asp">
<input name="action" type="hidden" value="ok">
  <label></label>
  <label></label>
  <table width="100%" border="0">
    <tr>
      <td>&nbsp;</td>
      <td><span class="STYLE1">������ɾ��������¼���޷��ָ�����С��ʹ��!�����¼�ǳ��������ĵȴ���<br>
        <br>
      </span></td>
    </tr>
    <tr>
      <td width="149"><div align="right">����ɾ��:</div></td>
      <td width="853">��
        <input type="text" name="startdate"   value="��ѡ������" onclick="SelectDate(this,'yyyy-MM-dd')" >
        <label>��</label>  <input type="text" name="enddate"   value="��ѡ������" onclick="SelectDate(this,'yyyy-MM-dd')" ></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input type="submit" name="Submit" value="����ɾ��������¼" onClick="return confirm('ȷ��Ҫɾ��������');"></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;&nbsp;</td>
    </tr>
  </table>
  <label><br>
  <br>
  </label>
</form>