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
    <td height=25 class=menu_title onmouseover=this.className='menu_title'; onmouseout=this.className='menu_title'; background="../images/membertitle.gif" id=menuTitle29 onClick="showsubmenu(29)" style="cursor:hand;"><span>�ҵ��˻�</span> </td>
  </tr>
  <tr>
    <td  id='submenu29'><div class=sec_menu style="width:218">
        <table cellpadding=0 cellspacing=0 align=center width=190>
<tr>
  <td height=20>��Ա��<font 
color="#ff0000"><%=session("useridname")%></font><%if vip="1" then%><img src="../images/jieducm_vip.gif"  alt="���VIP������ֵ:<%call vipxinyu(session("useridname"))%>" /><%end if%> ����</td>
</tr>
<tr> 
 <td height=20>��ӵ�У�<font color=#ff0000>
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
%></font> Ԫ</td></tr>
 <tr><td height=20>�����㣺<font color=#ff0000>
<%
if (fabudian)=0 then
response.Write("0.0")
elseif fabudian<1 then
fabudian=formatnumber((fabudian),2)
response.Write(fabudian)
else
fabudian=formatnumber((fabudian),2)
response.Write(fabudian)
end  if%></font> ��</td>
</tr>
 <tr><td height=20>�� �֣�<font 
color=#ff0000><%=jifei%></font>��</td>
</tr>
 <%if vip=1 then%>
 <tr><td height=20>VIP��Ч�ڣ�<FONT color=#ff0000><%=vipdate%></FONT>��</td>
</tr>
 <tr>
   <td height=20>����ʱ�䣺<%=FormatDate(vipdel,2)%></td>
 </tr>
 <tr><td height=20>VIP����ֵ��<FONT color=#ff0000><%call vipxinyu(session("useridname"))%></FONT></td>
</tr>
<%end if%>
</table>
    </div>
      </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=218 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="../images/membertitle.gif" id=menuTitle8 onClick="showsubmenu(274)" style="cursor:hand;"><span>�����㽻���г�</span> </td>
  </tr>
  <tr>
    <td  id='submenu274' ><div class=sec_menu style="width:218">
        <table cellpadding=0 cellspacing=0 align=center width=190  class=LeftNews>
          <tr>
            <td height=20> <a href="../user/paypoint.asp">���򷢲���</a></td>
          </tr>
          <tr>
            <td height=20><a href="../user/pay.asp"><FONT  color=red>���۷�����</FONT></a></td>
          </tr>
          
          <tr>
            <td height=20><a href="../user/paysell.asp?action=1"  >�����еķ�����</a></td>
          </tr>
          <tr>
            <td height=20><a href="../user/paysell.asp">�ѳ��۵ķ�����</a></td>
          </tr>
       
        </table>
      </div>
        </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=218 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="../images/membertitle.gif" id=menuTitle8 onClick="showsubmenu(243)" style="cursor:hand;"><span>�������Ϲ���</span> </td>
  </tr>
  <tr>
    <td  id='submenu243'><div class=sec_menu style="width:218">
        <table cellpadding=0 cellspacing=0 align=center width=190  class=LeftNews>
		<tr>
		  <td height=20><a href="Express.asp"><font color="#FF0000">���ɿ�ݵ���</font></a></td>
		  </tr>
		<tr>
            <td height=20><A  href="../user/Center_Userlock.asp" >���ֻ� </A></td>
          </tr>
		   <tr >
		   <td height=20><a  href="../user/SetIp.asp" >���õ�½IP </a></td>
		   </tr>
		 <tr >
            <td height=20><a  href="../user/sms.asp" >�����ֻ����� </a></td>
          </tr>
         
          <tr>
            <td height=20><A   href="../user/guest_info.asp" >�޸�����</A></td>
          </tr>
          <tr>
            <td height=20><a href="../user/SetPassword.asp" >����/������</a> </td>
          </tr>
          <tr>
            <td height=20><A   href="../union/index.asp" target="_blank"  >��Ҫ�ƹ�</A></td>
          </tr>
          <tr>
            <td height=20><A   href="../user/promotion.asp"   >�ҵ��ƹ�</A></td>
          </tr>
          <tr>
            <td height=20><a  href="../user/ListTop.asp">ˢ������</a></td>
          </tr>
          <tr>
            <td height=20><A   href="../user/user.asp" >վ����</A></td>
          </tr>
         
         
          <tr>
            <td height=20><a href="../user/login.asp"><font color="#FF0000">����VIP</font></a></td>
          </tr>
          <tr>
            <td height=20><A  onclick="return confirm('ȷ���˳�������');"  href="../user/exit.asp">
										<FONT  color=red>��ȫ�˳�</FONT></A></td>
          </tr>
        </table>
      </div>
        </td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=218 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="../images/membertitle.gif" id=menuTitle8 onClick="showsubmenu(24)" style="cursor:hand;"><span>�������</span> </td>
  </tr>
  <tr>
    <td  id='submenu24'><div class=sec_menu style="width:218">
        <table cellpadding=0 cellspacing=0 align=center width=190  class=LeftNews>
          <tr>
            <td height=20><a href="../user/md5_pay.asp">�ʺų�ֵ</a></td>
          </tr>
          <tr>
            <td height=20><a href="../user/MoneyOrPush.asp">��Ҫ�һ�</a></td>
          </tr>
          
          <tr>
            <td height=20><a href="../user/mai.asp">���򷢲���</a></td>
          </tr>
		  <tr>
            <td height=20><a href="../user/MoneyOrPush.asp">����������</a></td>
          </tr
		  ><tr>
            <td height=20><a href="../user/tuoguan.asp">��Ҫ�й�</a></td>
          </tr>
          <tr>
            <td height=20><a href="../user/ReMoney.asp">��Ҫ����</a></td>
          </tr>
          <tr>
            <td height=20><a href="../user/remoneylist.asp">���ּ�¼</a></td>
          </tr>
          <tr>
            <td height=20><a href="../user/record.asp">������¼</a></td>
          </tr>
        </table>
      </div>
        </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=218 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="../images/membertitle.gif" id=menuTitle8 onClick="showsubmenu(242)" style="cursor:hand;"><span>�������</span> </td>
  </tr>
  <tr>
    <td  id='submenu242'><div class=sec_menu style="width:218">
        <table cellpadding=0 cellspacing=0 align=center width=190  class=LeftNews>
          <tr>
            <td height=20><a href="../tribe/">�ҵĲ���</a></td>
          </tr>
          <tr>
            <td height=20><a href="../tribe/applytribe.asp">��������</a></td>
          </tr>
          
          <tr>
            <td height=20><a href="../tribe/manage.asp">������</a></td>
          </tr>
         
        </table>
      </div>
        </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=218 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="../images/membertitle.gif" id=menuTitle8 onClick="showsubmenu(241)" style="cursor:hand;"><span>��������</span> </td>
  </tr>
  <tr>
    <td  id='submenu241'><div class=sec_menu style="width:218">
          <table cellpadding=0 cellspacing=0 align=center width=190  class=LeftNews>
          <tr>
            <td height=20><a href="../user/tousu.asp">��Ҫ����</a></td>
          </tr>
          <tr>
            <td height=20><a href="../user/complaintre.asp">���ܵ�������</a></td>
          </tr>
          <tr>
            <td height=20><a href="../user/complaintmy.asp">�ҵ�����</a><font color="#0000FF">&nbsp;</font></td>
          </tr>
            <tr>
            <td height=20><a href="../user/complainth.asp">�ٷ�������</a><font color="#0000FF">&nbsp;</font></td>
          </tr>
		   <tr>
            <td height=20><a href="../user/name.asp">�ҵĺ�����</a><font color="#0000FF">&nbsp;</font></td>
          </tr>
		    <tr>
            <td height=20><a href="../user/comps.asp">�ٷ�����</a><font color="#0000FF">&nbsp;</font></td>
          </tr>
        </table>
      </div>
       </td>
  </tr>
</table>
</DIV>