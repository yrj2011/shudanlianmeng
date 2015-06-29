<%@language=vbscript codepage=936 %>
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：七  七 传 媒
'官方主页：http://www.778892.com
'技术支持：xsh
'QQ;859680966
'电话：15889679112
'版权：版权属于七七传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2013 七七传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
option explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
'主要是使随机出现的图片数字随机
%>
<html>
<head>
<title>管理员登录</title>
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
		alert("请输入用户名！");
		document.Login.UserName.focus();
		return false;
	}
	if(document.Login.Password.value == "")
	{
		alert("请输入密码！");
		document.Login.Password.focus();
		return false;
	}
	if (document.Login.CheckCode.value==""){
       alert ("请输入您的验证码！");
       document.Login.CheckCode.focus();
       return(false);
    }
}

function CheckBrowser() 
{
  var app=navigator.appName;
  var verStr=navigator.appVersion;
  if (app.indexOf('Netscape') != -1) {
    alert("MY动力友情提示：\n    你使用的是Netscape浏览器，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。");
  } 
  else if (app.indexOf('Microsoft') != -1) {
    if (verStr.indexOf("MSIE 3.0")!=-1 || verStr.indexOf("MSIE 4.0") != -1 || verStr.indexOf("MSIE 5.0") != -1 || verStr.indexOf("MSIE 5.1") != -1)
      alert("MY动力友情提示：\n    您的浏览器版本太低，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。");
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
     <font color="#999999">后台管理页面需要屏幕分辨率为</font> <font color="#FF9900"><strong>1024*768</strong></font> <font color="#999999">或以上才能达到最佳浏览效果！　　需要浏览器为</font> <font color="#FF9900"><strong>IE5.5</strong></font> <font color="#999999">或以上版本才能正常运行！！！</font>      
      </marquee></td>
    </tr>
    <tr>
      <td height="166" colspan="3"><img src="Images/fp.gif" width="378" height="166"></td>
      <td width="394" background="Images/admin_login_44.gif">
	    <table width="300" border="0" cellspacing="8" cellpadding="0">
          <tr> 
            <td align="right"><span class="Glow">用户名称：</span></td>
            <td colspan="2"><input name="UserName"  type="text"  id="UserName4" maxlength="20" style="width:160px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); "></td>
          </tr>
          <tr> 
            <td align="right"><span class="Glow">用户密码：</span></td>
            <td colspan="2"><input name="Password"  type="password" maxlength="20" style="width:160px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); "></td>
          </tr>
          <tr> 
            <td align="right"><span class="Glow">验 证 码：</span></td>
            <td><input name="CheckCode" size="6" maxlength="4" style="border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); ">
              <font color="#FFFFFF">请在左边输入</font>            <img src="inc/checkcode.asp"></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="3"> <div align="center"> 
                <input   type="submit" name="Submit" value=" 确&nbsp;认 " style="font-size: 9pt; height: 19; width: 60; color: #000000; background-color: #E1F4EE; border: 1 solid #E1F4EE" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#E1F4EE'">
                &nbsp; 
                <input name="reset" type="reset"  id="reset" value=" 清&nbsp;除 " style="font-size: 9pt; height: 19; width: 60; color: #000000; background-color: #E1F4EE; border: 1 solid #E1F4EE" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#E1F4EE'">
                <br>
              </div></td>
          </tr>
        </table>
	  </td>
      <td background="Images/admin_login_45.gif">&nbsp;</td>
    </tr>
    <tr>
      <td height="41" colspan="3" align="right" background="Images/admin_login_22.gif">Copyright&nbsp;&copy;&nbsp;2013 &nbsp;<a href="http://www.778892.com" target="_blank"><font color="#FF9900">七七传媒</font></a>&nbsp;.&nbsp;All&nbsp;Rights&nbsp;Reserved.</td>
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
