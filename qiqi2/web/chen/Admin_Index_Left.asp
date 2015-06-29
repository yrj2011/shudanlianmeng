<%@language=vbscript codepage=936 %>
<%
Const PurviewLevel=0
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<html>
<head>
<title>管理导航</title>
<style type=text/css>
body  { background:#555555; margin:0px; font-family: Verdana, Arial, sans-serif, 宋体; font-size: 9pt; text-decoration: none; color:#555555;
SCROLLBAR-FACE-COLOR: #cccccc;
SCROLLBAR-HIGHLIGHT-COLOR: #cccccc;
SCROLLBAR-SHADOW-COLOR: #cccccc;
SCROLLBAR-3DLIGHT-COLOR: #cccccc;
SCROLLBAR-ARROW-COLOR: #555555;
SCROLLBAR-TRACK-COLOR: #555555;
SCROLLBAR-DARKSHADOW-COLOR: #cccccc;}
table  { border:0px; }
td  { font:normal 12px; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px; color:#555555; text-decoration:none; }
a:hover  { color:#ff6600;text-decoration:underline; }
.sec_menu  { border-left:1px solid white; border-right:1px solid white; border-bottom:1px solid white; overflow:hidden; background:#eeeeee; }
.menu_title  { }
.menu_title span  { position:relative; top:2px; left:8px; color:#555555; font-weight:bold; }
.menu_title2  { }
.menu_title2 span  { position:relative; top:2px; left:8px; color:#ff6600; font-weight:bold; }
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

<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
 <table width=158 border="0" align=center cellpadding=0 cellspacing=0>
  <tr>
    <td valign=top>	  <img src="images/title.gif" border="0" alt="七七传媒"> </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/title_bg_quit.gif id=menuTitle0> 
          <span><a href="Admin_Index_Main.asp" target=main><b>管理首页</b></a> | <a href=Admin_logout.asp target=_top><b>退出</b></a></span> 
        </td>
  </tr>
  <tr>
    <td style="display:" id='submenu0'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20>用户名：<%=AdminName%></td></tr>
<tr><td height=20>权&nbsp;&nbsp;限：<%
		  select case AdminPurview
		  	case 1
				response.write "超级管理员"
			case 2
				response.write "<a href=Admin_ShowPurview.asp target=main>普通管理员</a>"
		  end select
		  call  total_jieducm()
		  %></td></tr>
</table>
</div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
</div>
	</td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_8.gif" id=menuTitle29 onClick="showsubmenu(29)" style="cursor:hand;"><span>系统设置</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu29'><div class=sec_menu style="width:158">
        <table cellpadding=0 cellspacing=0 align=center width=130>
             <%if AdminPurview=1 or AdminPurview_class1=1 then%>
			  <tr> 
                <td height=20><a href=Admin_SiteConfig.asp target=main>网站信息配置</a>  </td>
              </tr>
			  <%end if
			  if AdminPurview=1 or AdminPurview_class2=1 then%>
      
              <tr>
                <td height=20><A href="send_ms_message.asp" target="main">发送邮件</A></td>
              </tr>
              <tr>
                <td height=20><A href="send_zn_message.asp" target="main">发送站内信息</A></td>
              </tr>
                <tr><td height=20><A href="zn_message.asp" target="main">管理站内信息</A></td>
                </tr>
				<tr>
                  <td height=20><A href="phone_message.asp" target="main">管理手机短信</A></td>
                </tr>
				<%end if
				if AdminPurview=1 or AdminPurview_class3=1 then%>
<tr><td height=20><A href="Admin_FriendLinks.asp" target="main">友情连接</A></td>
</tr>
<%end if%>
<tr><td height=20><A href="qqadmin.asp" target="main">QQ客服管理</A></td>
</tr>
<%if AdminPurview=1 or AdminPurview_class4=1 then%>
<tr>
  <td height=20><A href="heiguan.asp" target="main">官方黑名单管理</A></td>
</tr>
<tr>
  <td height=20><A href="hei.asp" target="main">黑名单管理</A></td>
</tr>
<%end if
if AdminPurview=1 or AdminPurview_class5=1 then%>
<tr>
  <td height=20><A href="viewreturn.asp" target="main">留言管理</A></td>
</tr>
<%end if
if AdminPurview=1 or AdminPurview_class7=1 then%>
<tr>
  <td height=20><A href="exesql.asp" target="main">自定义执行SQL</A></td>
</tr>
<%end if
if AdminPurview=1 or AdminPurview_class8=1 then%>

<%end if%>
</table>
      </div>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=130>
            <tr>
              <td height=20></td>
            </tr>
          </table>
      </div></td>
  </tr>
</table>
<%if AdminPurview=1 or AdminPurview_class9=1 then%>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_1.gif" id=menuTitle1 onClick="showsubmenu(1)" style="cursor:hand;"> 
          <span>文章管理</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu1'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height="20"><a href=Admin_ArticleAdd1.asp target=main>添加文章</a> | <a href=Admin_ArticleManage.asp target=main>文章管理</a></td>
</tr>



<tr>
  <td height=20><a href=Admin_ArticleRecyclebin.asp target=main>回收站管理</a></td>
</tr>

<tr><td height=20><a href=Admin_Class_Article.asp?Action=Add target=main>文章栏目添加</a> | <a href=Admin_Class_Article.asp target=main>管理</a></td></tr>

</table>
</div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
</div>
	</td>
  </tr>
</table>
<%end if%>

<%if AdminPurview=1 or AdminPurview_class10=1 then%>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_8.gif" id=menuTitle8 onClick="showsubmenu(422)" style="cursor:hand;"><span>充值卡管理</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu422'><div class=sec_menu style="width:158">
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr>
            <td height=20><A href="Admin_Jicar.asp" target="main">充值卡管理</A></td>
          </tr>
          <tr>
            <td height=20><a href="Admin_Okjicar.asp" target="main">点卡充值记录</a></td>
          </tr>
          
           <tr>
            <td height=20><A href="Admin_Jiagesear.asp" target="main">点卡密码查询</A> </td>
          </tr>
        </table>
      </div>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=130>
            <tr>
              <td height=20></td>
            </tr>
          </table>
      </div></td>
  </tr>
</table>
<%end if
if AdminPurview=1 or AdminPurview_class11=1 then%>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_8.gif" id=menuTitle8 onClick="showsubmenu(203)" style="cursor:hand;"><span>任务清理</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu203'><div class=sec_menu style="width:158">
        <table cellpadding=0 cellspacing=0 align=center width=130>
		   <tr>
            <td height=20><a href=admin_delrecord.asp target=main >一键清理日志</a></td>
          </tr>
		   <tr>
            <td height=20><a href=admin_deltaobao.asp target=main >一键清理任务</a></td>
          </tr>
          <tr>
            <td height=20><a href=admin_Mission.asp target=main>清理任务</a></td>
          </tr>
        </table>
      </div>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=130>
            <tr>
              <td height=20></td>
            </tr>
          </table>
      </div></td>
  </tr>
</table>
<%end if
if AdminPurview=1 or AdminPurview_class12=1 then%>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_8.gif" id=menuTitle8 onClick="showsubmenu(276)" style="cursor:hand;"><span>充值提现任务管理</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu276'><div class=sec_menu style="width:158">
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr>
            <td height=20><a href=admin_MyMission4.asp target=main>充值提现任务管理</a></td>
          </tr>
		   <tr>
            <td height=20><a href=admin_MyMissionjie4.asp target=main>已接充值提现任务</a></td>
          </tr>
        </table>
      </div>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=130>
            <tr>
              <td height=20></td>
            </tr>
          </table>
      </div></td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_8.gif" id=menuTitle8 onClick="showsubmenu(2)" style="cursor:hand;"><span>淘宝区任务管理</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu2'><div class=sec_menu style="width:158">
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr>
            <td height=20><a href=admin_MyMission.asp target=main>淘宝任务管理</a></td>
          </tr>
		   <tr>
            <td height=20><a href=admin_MyMissionjie.asp target=main>已接淘宝任务管理</a></td>
          </tr>
        </table>
      </div>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=130>
            <tr>
              <td height=20></td>
            </tr>
          </table>
      </div></td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_8.gif" id=menuTitle8 onClick="showsubmenu(236)" style="cursor:hand;"><span>收藏任务管理</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu236'><div class=sec_menu style="width:158">
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr>
            <td height=20><a href=admin_MyMission5.asp target=main>收藏任务管理</a></td>
          </tr>
		   <tr>
            <td height=20><a href=admin_MyMissionjie5.asp target=main>已接收藏任务</a></td>
          </tr>
        </table>
      </div>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=130>
            <tr>
              <td height=20></td>
            </tr>
          </table>
      </div></td>
  </tr>
</table>
<%end if
if AdminPurview=1 or AdminPurview_class13=1 then%>


<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_8.gif" id=menuTitle8 onClick="showsubmenu(24)" style="cursor:hand;"><span>拍拍区任务管理</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu24'><div class=sec_menu style="width:158">
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr>
            <td height=20><a href=admin_MyMission1.asp target=main>拍拍任务管理</a></td>
          </tr>
		    <tr>
            <td height=20><a href=admin_MyMissionjie3.asp target=main>已接拍拍任务管理</a></td>
          </tr>
        </table>
      </div>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=130>
            <tr>
              <td height=20></td>
            </tr>
          </table>
      </div></td>
  </tr>
</table>
<%end if%>


<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_8.gif" id=menuTitle8 onClick="showsubmenu(42)" style="cursor:hand;"><span>用户管理</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu42'><div class=sec_menu style="width:158">
        <table cellpadding=0 cellspacing=0 align=center width=130>
		<%if AdminPurview=1 or AdminPurview_class16=1 then%>

          <tr>
            <td height=20><a href=usermanage.asp target=main>用户管理</a></td>
          </tr>
           <%end if%>
          
           <tr>
             <td height=20><a href=chengfaman.asp target=main>奖惩管理</a></td>
           </tr>
		   <%if AdminPurview=1 or AdminPurview_class17=1 then%>

           <tr>
             <td height=20><a href=admin_recordm.asp target=main>管理操作记录</a></td>
           </tr>
           <%end if
		   if AdminPurview=1 or AdminPurview_class18=1 then%>

           <tr>
            <td height=20><a href=admin_record.asp target=main>会员操作记录</a></td>
          </tr>
		  <%end if%>
        </table>
      </div>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=130>
            <tr>
              <td height=20></td>
            </tr>
          </table>
      </div></td>
  </tr>
</table>
<%if AdminPurview=1 or AdminPurview_class19=1 then%>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_2.gif" id=menuTitle2 onClick="showsubmenu(3)" style="cursor:hand;"> 
          <span>财务管理 </span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu3'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>
               <tr>
			    <td height=20><a href="jieducm_car.asp" target=main>网站点卡管理</a></td>
		      </tr>
			   <tr>
			     <td height=20><a href="jieducm_vipcar.asp" target=main>VIP卡管理</a></td>
		      </tr>
			                 <tr>
                 <td height=20><a href="jieducm_pay.asp" target="main">发布点交易市场</a></td>
               </tr>
			   <tr>
                 <td height=20><a href="AutoCZ_Manage.asp" target=main><font color="red">转账充值管理</font></a></td>
               </tr>
		      <tr>
			    <td height=20><a href="admin_bankroll.asp" target=main>提现记录</a></td>
		      </tr>
               <tr>
                 <td height=20><a href="chongrecord.asp" target=main>充值记录</a></td>
               </tr>
               <tr>
			    <td height=20><a href="jieducm_jiesuan.asp" target=main>平台余额结算</a></td>
		      </tr>
            </table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>
<%end if%>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_2.gif" id=menuTitle2 onClick="showsubmenu(323)" style="cursor:hand;"> 
          <span>部落管理 </span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu323'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>
              
			  <tr>
			    <td height=20><a href="tribemanage.asp" target=main>部落管理</a></td>
		      </tr>
               <tr>
			    <td height=20><a href="tribebook.asp" target=main>部落留言管理</a></td>
		      </tr>
            </table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>
<%if AdminPurview=1 then%>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_4.gif" id=menuTitle4 onClick="showsubmenu(4)" style="cursor:hand;"> 
          <span>管理员管理</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu4'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>

              <tr>
                <td height=20><a href=Admin_Admin.asp?Action=Add target=main>管理员添加</a> 
                  | <a href=Admin_Admin.asp target=main>管理</a></td>
              </tr>

            </table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>


<%end if%>
<%if AdminPurview=1 or AdminPurview_class20=1 then%>


<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_6.gif" id=menuTitle6 onClick="showsubmenu(6)" style="cursor:hand;"> 
          <span>广告管理</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu6'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr>
                <td height=20><a href="HL_admin_gg.asp" target="main">广告管理</a></td>
          </tr>
<tr>
                <td height=20><a href="HL_admin_gg.asp?action=tj" target="main">添加广告</a></td>
          </tr>
          <tr>
                <td height=20><a href="HL_admin_manwei.asp" target="main">广告位管理</a></td>
          </tr>
          <tr>
                <td height=20><a href="HL_admin_ggmanlei.asp" target="main">广告类别管理</a></td>
          </tr>
                    <tr>
                <td height=20><a href="guanggao.asp" target="main">首页FLASH广告</a></td>
          </tr>
</table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>
<%end if
if AdminPurview=1 or AdminPurview_class21=1 then%>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_6.gif" id=menuTitle6 onClick="showsubmenu(62)" style="cursor:hand;"> 
          <span>申诉管理</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu62'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr>
                <td height=20><a href="shenshu.asp" target="main">申诉管理</a></td>
          </tr>

</table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>

<%end if
if AdminPurview=1 or AdminPurview_class22=1 then%>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_6.gif" id=menuTitle6 onClick="showsubmenu(622)" style="cursor:hand;"> 
          <span>SQL注入管理</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu622'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr>
                <td height=20><a href="sqlInConfig.asp?moduleid=320" target="main">系统设置</a></td>
          </tr>
		  <tr>
                <td height=20><a href="sqlInList.asp?moduleid=321" target="main">查看信息</a></td>
          </tr>

</table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>
<%end if%>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_9.gif" id=menuTitle9> 
          <span>网站信息</span></td>
  </tr>
  <tr>
    <td>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20 valign="top"><br>
程序开发：<a href="http://www.688880.com" target="_blank">浩宇网络</a><br>
制作QQ：1662424999<br>
&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
</table>
	  </div>
	</td>
  </tr>
</table>
</body>
</html>