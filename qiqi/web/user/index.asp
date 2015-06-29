<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="checklogin.asp"-->
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
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">				
<head>
<title>会员中心-<%=stupname%> -<%=stuptitle%></title>
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
        <td height="30" align="left" background="../images/zc-bg.gif"><img src="../images/icon13.gif" width="15" height="15" /> <%=stupname%>&gt;&gt;帐号信息：</td>
        </tr>
      <tr>
        <td height="500" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0"  id="con_three_1">
  
            <tr>
              <td height="55" align="center" class="borderc"><IMG src="../images/c3.gif">你好： </td>
              <td width="23%" align="left" class="borderc"><Font color="#FF0000"><%=session("useridname")%></Font><%if vip="1" then%><img src="../images/jieducm_vip.gif"  alt="尊贵VIP" /><%end if%>  &nbsp;<%call tribename(tribeid)%></td>
              <td width="63%" align="left" class="borderc">&nbsp;
			  
			  <a href="md5_pay.asp"><img src="../images/jieducm_congzhi.jpg" width="140" height="40" border="0" /></a>&nbsp;&nbsp;&nbsp;</td>
            </tr>
            <tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc">你本次登陆的IP城市是：
                <script type="text/javascript" charset=gb2312>   
document.write(IPData[2]);document.write(IPData[3]);   
</script>  </td>
            </tr>
            <tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc">你上次登陆的时间是：<font color="#FF0000"><%=nowe%></font>，如果你在这个时间没有登陆过，请联系客服。 </td>
              </tr>
            <tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc"><IMG src="../images/c8.gif">你的存款：<font color="#FF0000">
              
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
%></font> 元；
发布点：<font color="#FF0000">
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
</font>点； 刷客积分：
<%call xinyu(jifei)%></td>
              </tr>
            <%if vip=1 then%><tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc"> 
VIP有效期：<FONT color=#ff0000><%=vipdate%></FONT>天； 加入时间：<%=FormatDate(vipdel,2)%>；  VIP信誉值：<FONT color=#ff0000><%call vipxinyu(session("useridname"))%></FONT>
<br></td>
            </tr><%end if%>
            <tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc"><font color="#FF0000"><font color="#FF0000"></font>&nbsp; </font> <a href="record.asp">查看交易记录</a> </td>
            </tr>
            
            <tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
           
              <td colspan="2" align="left" class="borderc">你现在有<A href="user.asp"><font color="#FF0000"><%=weidu1%></font> 封系统站内信</A>没有阅读，<A href="users.asp"><font color="#FF0000"><%=weidu%> </font>封私人站内信</A>没有阅读 </td>
            </tr>
            <tr>
              <td width="14%" height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc">你现在收到<A href="complaintre.asp"><font color="#FF0000"><%=shengsu%></font>个</A>被申诉，请尽快处理</td>
              </tr>
                  <tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc">买家操作： 等待我付款（<a href="../taobao/ReMission.asp?sort=1"><font color="#FF0000"><%call jieducm_msg(1)%></font></a>）  等待卖家发货（<a href="../taobao/ReMission.asp?sort=2"><font color="#FF0000"><%call jieducm_msg(2)%></font></a>） 等待我收货好评（<a href="../taobao/ReMission.asp?sort=3"><font color="#FF0000"><%call jieducm_msg(3)%></font></a>）<br />
                卖家操作： 等待接手（<a href="../taobao/MyMission.asp?sort=0"><font color="#FF0000"><%call jieducm_msg(4)%></font></a>）等待审核（<a href="../taobao/MyMission.asp?sort=1"><font color="#FF0000">
                <%call jieducm_msg(8)%>
                </font></a>）    等待我发货（<a href="../taobao/MyMission.asp?sort=2"><font color="#FF0000">
                <%call jieducm_msg(5)%></font></a>）   等待买家确认（<a href="../taobao/MyMission.asp?sort=3"><font color="#FF0000"><%call jieducm_msg(6)%></font></a>）     等待我核查好评（<a href="../taobao/MyMission.asp?sort=4"><font color="#FF0000"><%call jieducm_msg(7)%></font></a>）</td>
            </tr>
			
			 <tr>
              <td height="55" align="center" class="borderc">&nbsp;</td>
              <td colspan="2" align="left" class="borderc">平台提示：<font color="#FF0000">请尽快核对货款，按卖家规定时间完成平台任务，超过规定时间12小时后，会员有权进行申诉处理！</font><br /></td>
            </tr>
          </table>
  <!--2项 --></td>
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