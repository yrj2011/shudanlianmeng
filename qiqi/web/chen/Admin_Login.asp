<%@language=vbscript codepage=936 %>
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
option explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
'��Ҫ��ʹ������ֵ�ͼƬ�������
%>
<html>
<head>
<title>����Ա��¼</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Admin_Style.css">
<script language=javascript>
<!--
function SetFocus()
{
if (document.Login.UserName.value=="")
	document.Login.UserName.focus();
else
	document.Login.UserName.select();
}
function CheckForm()
{
	if(document.Login.UserName.value=="")
	{
		alert("�������û�����");
		document.Login.UserName.focus();
		return false;
	}
	if(document.Login.Password.value == "")
	{
		alert("���������룡");
		document.Login.Password.focus();
		return false;
	}
	if (document.Login.CheckCode.value==""){
       alert ("������������֤�룡");
       document.Login.CheckCode.focus();
       return(false);
    }
}

function CheckBrowser() 
{
  var app=navigator.appName;
  var verStr=navigator.appVersion;
  if (app.indexOf('Netscape') != -1) {
    alert("MY����������ʾ��\n    ��ʹ�õ���Netscape����������ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��");
  } 
  else if (app.indexOf('Microsoft') != -1) {
    if (verStr.indexOf("MSIE 3.0")!=-1 || verStr.indexOf("MSIE 4.0") != -1 || verStr.indexOf("MSIE 5.0") != -1 || verStr.indexOf("MSIE 5.1") != -1)
      alert("MY����������ʾ��\n    ����������汾̫�ͣ����ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��");
  }
}
//-->
</script>
</head>
<body bgcolor="#000000"  leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<p>&nbsp;</p>
<p>&nbsp;</p>
<p><form name="Login" action="Admin_ChkLogin.asp" method="post" target="_parent" onSubmit="return CheckForm();">
  <table width="100%"  border="0" align="left" cellpadding="0" cellspacing="0">
    <tr>
      <td width="125" height="39" background="Images/admin_login_2.gif"><img src="Images/admin_login_1.gif" width="125" height="39"></td>
      <td width="64" background="Images/admin_login_2.gif">&nbsp;</td>
      <td width="189" background="Images/admin_login_2.gif"><img src="Images/admin_login_3.gif" width="189" height="39"></td>
      <td colspan="2" valign="bottom" background="Images/admin_login_2.gif"><MARQUEE scrollAmount=1 scrollDelay=4 width=300 align="left" onMouseOver="this.stop()" onMouseOut="this.start()">
     <font color="#999999">��̨����ҳ����Ҫ��Ļ�ֱ���Ϊ</font> <font color="#FF9900"><strong>1024*768</strong></font> <font color="#999999">�����ϲ��ܴﵽ������Ч����������Ҫ�����Ϊ</font> <font color="#FF9900"><strong>IE5.5</strong></font> <font color="#999999">�����ϰ汾�����������У�����</font>      
      </marquee></td>
    </tr>
    <tr>
      <td height="166" colspan="3"><img src="Images/fp.gif" width="378" height="166"></td>
      <td width="394" background="Images/admin_login_44.gif">
	    <table width="300" border="0" cellspacing="8" cellpadding="0">
          <tr> 
            <td align="right"><span class="Glow">�û����ƣ�</span></td>
            <td colspan="2"><input name="UserName"  type="text"  id="UserName4" maxlength="20" style="width:160px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); "></td>
          </tr>
          <tr> 
            <td align="right"><span class="Glow">�û����룺</span></td>
            <td colspan="2"><input name="Password"  type="password" maxlength="20" style="width:160px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); "></td>
          </tr>
          <tr> 
            <td align="right"><span class="Glow">�� ֤ �룺</span></td>
            <td><input name="CheckCode" size="6" maxlength="4" style="border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); ">
              <font color="#FFFFFF">�����������</font>            <img src="inc/checkcode.asp"></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="3"> <div align="center"> 
                <input   type="submit" name="Submit" value=" ȷ&nbsp;�� " style="font-size: 9pt; height: 19; width: 60; color: #000000; background-color: #E1F4EE; border: 1 solid #E1F4EE" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#E1F4EE'">
                &nbsp; 
                <input name="reset" type="reset"  id="reset" value=" ��&nbsp;�� " style="font-size: 9pt; height: 19; width: 60; color: #000000; background-color: #E1F4EE; border: 1 solid #E1F4EE" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#E1F4EE'">
                <br>
              </div></td>
          </tr>
        </table>
	  </td>
      <td background="Images/admin_login_45.gif">&nbsp;</td>
    </tr>
    <tr>
      <td height="41" colspan="3" align="right" background="Images/admin_login_22.gif">Copyright&nbsp;&copy;&nbsp;2013 &nbsp;<a href="http://www.778892.com" target="_blank"><font color="#FF9900">���ߴ�ý</font></a>&nbsp;.&nbsp;All&nbsp;Rights&nbsp;Reserved.</td>
      <td colspan="2" align="right" background="Images/admin_login_22.gif"><img src="Images/admin_login_5.gif" width="125" height="41"></td>
    </tr>
  </table>
</p> <br>   
</form>
<script language="JavaScript" type="text/JavaScript">
CheckBrowser();
SetFocus(); 
</script>
</body>
</html>
