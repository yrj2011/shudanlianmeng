<style type=text/css>
.sec_menu  { border-left:1px solid white; border-right:1px solid white; border-bottom:1px solid white; overflow:hidden;  }
.menu_title  { }
.menu_title span  {
	position:relative;
	top:2px;
	left:28px;
	color:#ffffff;
	font-weight:bold;
}
.menu_title2 span  { position:relative; top:2px; left:28px; color:#ffffff; font-weight:bold; }
</style>
<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
</SCRIPT>
<DIV 
      style="BORDER-RIGHT: #8fc2d1 1px solid; BORDER-TOP: #8fc2d1 1px solid; BACKGROUND: url(/image/task_bg_02.jpg) #ffffff repeat-x 50% top; FLOAT: left; BORDER-LEFT: #8fc2d1 1px solid; WIDTH: 220px; BORDER-BOTTOM: #8fc2d1 1px solid">
<table cellpadding=0 cellspacing=0 width=218 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title'; onmouseout=this.className='menu_title'; background="../images/membertitle.gif" id=menuTitle29 onClick="showsubmenu(29)" style="cursor:hand;"><span>我的账户</span> </td>
  </tr>
  <tr>
    <td  id='submenu29'><div class=sec_menu style="width:218">
        <table cellpadding=0 cellspacing=0 align=center width=190>
<tr>
  <td height=20>会员：<font 
color="#ff0000"><%=session("useridname")%></font><%if vip="1" then%><img src="../images/jieducm_vip.gif"  alt="尊贵VIP，信誉值:<%call vipxinyu(session("useridname"))%>" /><%end if%> 您好</td>
</tr>
<tr> 
 <td height=20>您拥有：<font color=#ff0000>
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
%></font> 元</td></tr>
 <tr><td height=20>发布点：<font color=#ff0000>
<%
if (fabudian)=0 then
response.Write("0.0")
elseif fabudian<1 then
fabudian=formatnumber((fabudian),2)
response.Write(fabudian)
else
fabudian=formatnumber((fabudian),2)
response.Write(fabudian)
end  if%></font> 点</td>
</tr>
 <tr><td height=20>积 分：<font 
color=#ff0000><%=jifei%></font>点</td>
</tr>
 <%if vip=1 then%>
 <tr><td height=20>VIP有效期：<FONT color=#ff0000><%=vipdate%></FONT>天</td>
</tr>
 <tr>
   <td height=20>加入时间：<%=FormatDate(vipdel,2)%></td>
 </tr>
 <tr><td height=20>VIP信誉值：<FONT color=#ff0000><%call vipxinyu(session("useridname"))%></FONT></td>
</tr>
<%end if%>
</table>
    </div>
      </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=218 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="../images/membertitle.gif" id=menuTitle8 onClick="showsubmenu(274)" style="cursor:hand;"><span>发布点交易市场</span> </td>
  </tr>
  <tr>
    <td  id='submenu274' ><div class=sec_menu style="width:218">
        <table cellpadding=0 cellspacing=0 align=center width=190  class=LeftNews>
          <tr>
            <td height=20> <a href="../user/paypoint.asp">购买发布点</a></td>
          </tr>
          <tr>
            <td height=20><a href="../user/pay.asp"><FONT  color=red>出售发布点</FONT></a></td>
          </tr>
          
          <tr>
            <td height=20><a href="../user/paysell.asp?action=1"  >出售中的发布点</a></td>
          </tr>
          <tr>
            <td height=20><a href="../user/paysell.asp">已出售的发布点</a></td>
          </tr>
       
        </table>
      </div>
        </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=218 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="../images/membertitle.gif" id=menuTitle8 onClick="showsubmenu(243)" style="cursor:hand;"><span>个人资料管理</span> </td>
  </tr>
  <tr>
    <td  id='submenu243'><div class=sec_menu style="width:218">
        <table cellpadding=0 cellspacing=0 align=center width=190  class=LeftNews>
		<tr>
		  <td height=20><a href="Express.asp"><font color="#FF0000">生成快递单号</font></a></td>
		  </tr>
		<tr>
            <td height=20><A  href="../user/Center_Userlock.asp" >绑定手机 </A></td>
          </tr>
		   <tr >
		   <td height=20><a  href="../user/SetIp.asp" >设置登陆IP </a></td>
		   </tr>
		 <tr >
            <td height=20><a  href="../user/sms.asp" >发送手机短信 </a></td>
          </tr>
         
          <tr>
            <td height=20><A   href="../user/guest_info.asp" >修改资料</A></td>
          </tr>
          <tr>
            <td height=20><a href="../user/SetPassword.asp" >密码/操作码</a> </td>
          </tr>
          <tr>
            <td height=20><A   href="../union/index.asp" target="_blank"  >我要推广</A></td>
          </tr>
          <tr>
            <td height=20><A   href="../user/promotion.asp"   >我的推广</A></td>
          </tr>
          <tr>
            <td height=20><a  href="../user/ListTop.asp">刷客排行</a></td>
          </tr>
          <tr>
            <td height=20><A   href="../user/user.asp" >站内信</A></td>
          </tr>
         
         
          <tr>
            <td height=20><a href="../user/login.asp"><font color="#FF0000">加入VIP</font></a></td>
          </tr>
          <tr>
            <td height=20><A  onclick="return confirm('确定退出操作吗？');"  href="../user/exit.asp">
										<FONT  color=red>安全退出</FONT></A></td>
          </tr>
        </table>
      </div>
        </td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=218 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="../images/membertitle.gif" id=menuTitle8 onClick="showsubmenu(24)" style="cursor:hand;"><span>财务管理</span> </td>
  </tr>
  <tr>
    <td  id='submenu24'><div class=sec_menu style="width:218">
        <table cellpadding=0 cellspacing=0 align=center width=190  class=LeftNews>
          <tr>
            <td height=20><a href="../user/md5_pay.asp">帐号充值</a></td>
          </tr>
          <tr>
            <td height=20><a href="../user/MoneyOrPush.asp">我要兑换</a></td>
          </tr>
          
          <tr>
            <td height=20><a href="../user/mai.asp">购买发布点</a></td>
          </tr>
		  <tr>
            <td height=20><a href="../user/MoneyOrPush.asp">互赠发布点</a></td>
          </tr
		  ><tr>
            <td height=20><a href="../user/tuoguan.asp">我要托管</a></td>
          </tr>
          <tr>
            <td height=20><a href="../user/ReMoney.asp">我要提现</a></td>
          </tr>
          <tr>
            <td height=20><a href="../user/remoneylist.asp">提现记录</a></td>
          </tr>
          <tr>
            <td height=20><a href="../user/record.asp">操作记录</a></td>
          </tr>
        </table>
      </div>
        </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=218 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="../images/membertitle.gif" id=menuTitle8 onClick="showsubmenu(242)" style="cursor:hand;"><span>部落管理</span> </td>
  </tr>
  <tr>
    <td  id='submenu242'><div class=sec_menu style="width:218">
        <table cellpadding=0 cellspacing=0 align=center width=190  class=LeftNews>
          <tr>
            <td height=20><a href="../tribe/">我的部落</a></td>
          </tr>
          <tr>
            <td height=20><a href="../tribe/applytribe.asp">建立部落</a></td>
          </tr>
          
          <tr>
            <td height=20><a href="../tribe/manage.asp">管理部落</a></td>
          </tr>
         
        </table>
      </div>
        </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=218 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="../images/membertitle.gif" id=menuTitle8 onClick="showsubmenu(241)" style="cursor:hand;"><span>申诉中心</span> </td>
  </tr>
  <tr>
    <td  id='submenu241'><div class=sec_menu style="width:218">
          <table cellpadding=0 cellspacing=0 align=center width=190  class=LeftNews>
          <tr>
            <td height=20><a href="../user/tousu.asp">我要申诉</a></td>
          </tr>
          <tr>
            <td height=20><a href="../user/complaintre.asp">我受到的申诉</a></td>
          </tr>
          <tr>
            <td height=20><a href="../user/complaintmy.asp">我的申诉</a><font color="#0000FF">&nbsp;</font></td>
          </tr>
            <tr>
            <td height=20><a href="../user/complainth.asp">官方黑名单</a><font color="#0000FF">&nbsp;</font></td>
          </tr>
		   <tr>
            <td height=20><a href="../user/name.asp">我的黑名单</a><font color="#0000FF">&nbsp;</font></td>
          </tr>
		    <tr>
            <td height=20><a href="../user/comps.asp">官方奖罚</a><font color="#0000FF">&nbsp;</font></td>
          </tr>
        </table>
      </div>
       </td>
  </tr>
</table>
</DIV>