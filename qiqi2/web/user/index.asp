<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="checklogin.asp"-->
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷�����  �� �� ý
'�ٷ���ҳ��http://www.778892.com
'����֧�֣�xsh
'QQ;859680966
'�绰��15889679112
'��Ȩ����Ȩ�������ߴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2013 ���ߴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">				
<head>
<title>��Ա����-<%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<script src="http://fw.qq.com/ipaddress" type="text/javascript" charset=gb2312></script>   
</head>
<body>

<%
call check_jieducm_name(session("useridname"))	
if weidu1<>0 or weidu<>0 then
end if
%>
<!--#include file=../inc/jieducm_top.asp-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" >
<tr>
	    <td>
        <table width="910"  border="0" cellspacing="0" cellpadding="0">
	          <tr>
	          
	            <td width="140" valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"> </td>
    </tr>
</table>
<table id="menuToolBar" width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
    <td valign="top" bgcolor="FFFFFF">
    <!--#include file="userleft.asp"--></TD>
</TR>
</table></td>
	            <td width="15">&nbsp;</td>
	          
	            <td valign="top">
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"></td>
    </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#fefcde" class="bordera">
      <tr>
        <td height="30" align="left" background="../images/zc-bg.gif"><img src="../images/icon13.gif" width="15" height="15" /> <%=stupname%>&gt;&gt;�ʺ���Ϣ��</td>
        </tr>
      <tr>
        <td height="500" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0"  id="con_three_1">
  
            <tr>
              <td height="55" align="center" class="borderc"><IMG src="../images/c3.gif">��ã� </td>
              <td width="23%" align="left" class="borderc"><Font color="#FF0000"><%=session("useridname")%></Font><%if vip="1" then%><img src="../images/jieducm_vip.gif"  alt="���VIP" /><%end if%>  &nbsp;<%call tribename(tribeid)%></td>
              <td width="63%" align="left" class="borderc">&nbsp;
			  
			  <a href="md5_pay.asp"><img src="../images/jieducm_congzhi.jpg" width="140" height="40" border="0" /></a>&nbsp;&nbsp;&nbsp;</td>
            </tr>
            <tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc">�㱾�ε�½��IP�����ǣ�
                <script type="text/javascript" charset=gb2312>   
document.write(IPData[2]);document.write(IPData[3]);   
</script>  </td>
            </tr>
            <tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc">���ϴε�½��ʱ���ǣ�<font color="#FF0000"><%=nowe%></font>������������ʱ��û�е�½��������ϵ�ͷ��� </td>
              </tr>
            <tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc"><IMG src="../images/c8.gif">��Ĵ�<font color="#FF0000">
              
<%
if (cunkuan)=0 then
response.Write("0.00")
elseif cunkuan<1 then
cunkuan=formatnumber((cunkuan),2)
response.Write(cunkuan)
else
cunkuan=formatnumber((cunkuan),2)
response.Write(cunkuan)
end if
%></font> Ԫ��
�����㣺<font color="#FF0000">
<%
if (fabudian)=0 then
response.Write("0.0")
elseif fabudian<1 then
fabudian=formatnumber((fabudian),2)
response.Write(fabudian)
else
fabudian=formatnumber((fabudian),2)
response.Write(fabudian)
end  if%>
</font>�㣻 ˢ�ͻ��֣�
<%call xinyu(jifei)%></td>
              </tr>
            <%if vip=1 then%><tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc"> 
VIP��Ч�ڣ�<FONT color=#ff0000><%=vipdate%></FONT>�죻 ����ʱ�䣺<%=FormatDate(vipdel,2)%>��  VIP����ֵ��<FONT color=#ff0000><%call vipxinyu(session("useridname"))%></FONT>
<br></td>
            </tr><%end if%>
            <tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc"><font color="#FF0000"><font color="#FF0000"></font>&nbsp; </font> <a href="record.asp">�鿴���׼�¼</a> </td>
            </tr>
            
            <tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
           
              <td colspan="2" align="left" class="borderc">��������<A href="user.asp"><font color="#FF0000"><%=weidu1%></font> ��ϵͳվ����</A>û���Ķ���<A href="users.asp"><font color="#FF0000"><%=weidu%> </font>��˽��վ����</A>û���Ķ� </td>
            </tr>
            <tr>
              <td width="14%" height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc">�������յ�<A href="complaintre.asp"><font color="#FF0000"><%=shengsu%></font>��</A>�����ߣ��뾡�촦��</td>
              </tr>
                  <tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc">��Ҳ����� �ȴ��Ҹ��<a href="../taobao/ReMission.asp?sort=1"><font color="#FF0000"><%call jieducm_msg(1)%></font></a>��  �ȴ����ҷ�����<a href="../taobao/ReMission.asp?sort=2"><font color="#FF0000"><%call jieducm_msg(2)%></font></a>�� �ȴ����ջ�������<a href="../taobao/ReMission.asp?sort=3"><font color="#FF0000"><%call jieducm_msg(3)%></font></a>��<br />
                ���Ҳ����� �ȴ����֣�<a href="../taobao/MyMission.asp?sort=0"><font color="#FF0000"><%call jieducm_msg(4)%></font></a>���ȴ���ˣ�<a href="../taobao/MyMission.asp?sort=1"><font color="#FF0000">
                <%call jieducm_msg(8)%>
                </font></a>��    �ȴ��ҷ�����<a href="../taobao/MyMission.asp?sort=2"><font color="#FF0000">
                <%call jieducm_msg(5)%></font></a>��   �ȴ����ȷ�ϣ�<a href="../taobao/MyMission.asp?sort=3"><font color="#FF0000"><%call jieducm_msg(6)%></font></a>��     �ȴ��Һ˲������<a href="../taobao/MyMission.asp?sort=4"><font color="#FF0000"><%call jieducm_msg(7)%></font></a>��</td>
            </tr>
			
			 <tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc">ƽ̨��ʾ��<font color="#FF0000">�뾡��˶Ի�������ҹ涨ʱ�����ƽ̨���񣬳����涨ʱ��12Сʱ�󣬻�Ա��Ȩ�������ߴ���</font><br /></td>
            </tr>
          </table>
  <!--2�� --></td>
      </tr>
       
    </table></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

  <!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
  <%
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing
  %>
</body>
</html>