<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<%
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../jieducm/#jieducm.asp")
conn.Open connstr

if request("action")="save" then
	N_In		=replace(request.form("N_In"),"","''")
	Kill_IP		=request.form("Kill_IP")			
	WriteSql	=request.form("WriteSql")		
	alert_url	=request.form("alert_url")
	alert_info	=request.form("alert_info")
	kill_info	=request.form("kill_info")
	N_type		=request.form("N_type")
	Sec_Forms	=request.form("Sec_Forms")
	Sec_Form_open=request.form("Sec_Form_open")
	
	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From RC_sqlInconfig " ,Conn,3,3  
	    if not (rs.eof) then		
		rs("N_In")=N_In
		rs("Kill_IP")=Kill_IP
		rs("WriteSql")=WriteSql
		rs("alert_url")=alert_url
		rs("alert_info")=alert_info
		rs("kill_info")=kill_info
		rs("N_type")=N_type
		rs("Sec_Forms")=Sec_Forms
		rs("Sec_Form_open")=Sec_Form_open
    	rs.update
    	rs.close
	    end if

	response.write "<script>alert('���óɹ�');</script>"
end if

Set rsinfo=conn.execute("select * from RC_sqlInconfig")
	N_In		= rsinfo("N_In")
	Kill_IP		= rsinfo("Kill_IP")			
	WriteSql	= rsinfo("WriteSql")		
	alert_url	= rsinfo("alert_url")
	alert_info	= rsinfo("alert_info")
	kill_info	= rsinfo("kill_info")
	N_type		= rsinfo("N_type")
	Sec_Forms	= rsinfo("Sec_Forms")
	Sec_Form_open = rsinfo("Sec_Form_open")
	rsinfo.close
Set rsinfo=Nothing 
%>
<HTML>
<HEAD>
<TITLE>��ӭʹ����վ����ϵͳ</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
</HEAD>
<link href="../css/style.css" rel="stylesheet" type="text/css">
<BODY >
<table width="100" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="10"></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="553" height="22" background="../images/right_bg1.jpg">&nbsp;��ǰλ�ã�<%= lmwz%></td>
    <td width="436"><div align="right">[ <a href="javascript:history.go(-1)"><font color="#FF0000">����</font></a> ]</div></td>
  </tr>
</table>
<table width="100" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="5"></td>
  </tr>
</table>
<table width="98%" height="367" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#99D7EE" >
<form name="form" method="post" action="?moduleid=<%=moduleid%>&action=save">
  <tr>
    <td height="27" colspan=4 align="center" class="bb02">ϵͳ����</td>
  </tr>
  <tr >
    <td height="25" align=left bgcolor="#FFFFFF" class="r_bg">��Ҫ���˵Ĺؼ��֣�</td>
                <td width="80%" colspan="3" align=left bgcolor="#FFFFFF" class="right_bg"><input name="N_In" type="text" class="input1" id="r_str" style=" " value="<%=N_In%>" size="50">
        ��&quot;|&quot;�ֿ�</td>
  </tr>
               
              <tr > 
                <td width="20%" height="25" align=left bgcolor="#FFFFFF" class="r_bg" >�Ƿ��¼��������Ϣ��</td>
                <td colspan="3" align=left bgcolor="#FFFFFF" class="right_bg"><select name="WriteSql" id="r_kill">
          <option value="1" <%if WriteSql=1 Then response.write "selected"%>>��</option>
          <option value="0" <%if WriteSql=0 Then response.write "selected"%>>��</option>
        </select></td>
              </tr>
               
              <tr> 
                <td width="20%" height="25" align=left bgcolor="#FFFFFF" class="r_bg" >�Ƿ���������IP��</td>
                <td colspan="3" align=left bgcolor="#FFFFFF" class="right_bg"><select name="Kill_IP" id="r_kill">
          <option value="1" <%if Kill_IP=1 Then response.write "selected"%>>��</option>
          <option value="0" <%if Kill_IP=0 Then response.write "selected"%>>��</option>
        </select></td>
              </tr>
               
              <tr > 
                <td width="20%" height="25" align=left bgcolor="#FFFFFF" class="r_bg" >�Ƿ����ð�ȫҳ�棺</td>
                <td colspan="3" align=left bgcolor="#FFFFFF" class="right_bg"> <select name="Sec_Form_open" id="r_kill">
          <option value="1" <%if Sec_Form_open=1 Then response.write "selected"%>>��</option>
          <option value="0" <%if Sec_Form_open=0 Then response.write "selected"%>>��</option>
        </select>
		����������ܣ��������ȷ�ϴ�ҳ��������ˣ���ȷ���԰�ȫûӰ�죡</td>
              </tr>
               
              <tr> 
                <td width="20%" height="25" align=left bgcolor="#FFFFFF" class="r_bg" >����Ϊ��ȫ��ҳ�棺</td>
                <td colspan="3" align=left bgcolor="#FFFFFF" class="right_bg"><input name="Sec_Forms" type="text" class="input1" id="Sec_Forms" style=" " value="<%=Sec_Forms%>" size="50">
��&quot;|&quot;�ֿ�</td>
              </tr>
              <tr >
                <td height="25" align=left bgcolor="#FFFFFF" class="r_bg" >�����Ĵ���ʽ��</td>
                <td colspan="3" align=left bgcolor="#FFFFFF" class="right_bg"><select name="N_type" id="select">
                  <option value="1" <%if N_type=1 Then response.write "selected"%>>ֱ�ӹر���ҳ</option>
                  <option value="2" <%if N_type=2 Then response.write "selected"%>>�����ر�</option>
                  <option value="3" <%if N_type=3 Then response.write "selected"%>>��ת��ָ��ҳ��</option>
                  <option value="4" <%if N_type=4 Then response.write "selected"%>>�������ת</option>
                </select></td>
              </tr>
              <tr >
                <td height="25" align=left bgcolor="#FFFFFF" class="r_bg" >�������תUrl��</td>
                <td colspan="3" align=left bgcolor="#FFFFFF" class="right_bg"><input name="alert_url" type="text" class="input1" id="alert_url" style=" " value="http://www.778892.com/register.asp" size="30">
                ע�⣬����Ķ��ǰ�Ƿ��ţ�����Ӣ�ĵģ� </td>
              </tr>
              <tr >
                <td height="25" align=left bgcolor="#FFFFFF" class="r_bg" >������ʾ��Ϣ��</td>
                <td colspan="3" align=left bgcolor="#FFFFFF" class="right_bg"><textarea name="alert_info" cols="45" rows="4" id="r_err2" style=";  "><%=alert_info%></textarea>&nbsp;</td>
              </tr>
              <tr >
                <td height="25" align=left bgcolor="#FFFFFF" class="r_bg" >��ֹ������ʾ��Ϣ��</td>
                <td colspan="3" align=left bgcolor="#FFFFFF" class="right_bg"><textarea name="kill_info" cols="45" rows="4" id="r_err2" style=";  "><%=kill_info%></textarea></font></td>
              </tr>
              

              <tr > 
                <td width="20%" height="25" align=left bgcolor="#FFFFFF" class="r_bg" >&nbsp;</td>
                <td colspan="3" align=left bgcolor="#FFFFFF" class="right_bg"><input name="Submit" type="submit" class="submit" value=" �� �� "></td>
  </tr>
  </form>
</table>
<table width="100" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="10"></td>
  </tr>
</table>
</BODY>
</HTML>





