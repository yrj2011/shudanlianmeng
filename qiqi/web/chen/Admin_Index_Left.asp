<%@language=vbscript codepage=936 %>
<%
Const PurviewLevel=0
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<html>
<head>
<title>������</title>
<style type=text/css>
body  { background:#555555; margin:0px; font-family: Verdana, Arial, sans-serif, ����; font-size: 9pt; text-decoration: none; color:#555555;
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
    <td valign=top>	  <img src="images/title.gif" border="0" alt="���ߴ�ý"> </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/title_bg_quit.gif id=menuTitle0> 
          <span><a href="Admin_Index_Main.asp" target=main><b>������ҳ</b></a> | <a href=Admin_logout.asp target=_top><b>�˳�</b></a></span> 
        </td>
  </tr>
  <tr>
    <td style="display:" id='submenu0'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20>�û�����<%=AdminName%></td></tr>
<tr><td height=20>Ȩ&nbsp;&nbsp;�ޣ�<%
		  select case AdminPurview
		  	case 1
				response.write "��������Ա"
			case 2
				response.write "<a href=Admin_ShowPurview.asp target=main>��ͨ����Ա</a>"
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
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_8.gif" id=menuTitle29 onClick="showsubmenu(29)" style="cursor:hand;"><span>ϵͳ����</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu29'><div class=sec_menu style="width:158">
        <table cellpadding=0 cellspacing=0 align=center width=130>
             <%if AdminPurview=1 or AdminPurview_class1=1 then%>
			  <tr> 
                <td height=20><a href=Admin_SiteConfig.asp target=main>��վ��Ϣ����</a>  </td>
              </tr>
			  <%end if
			  if AdminPurview=1 or AdminPurview_class2=1 then%>
      
              <tr>
                <td height=20><A href="send_ms_message.asp" target="main">�����ʼ�</A></td>
              </tr>
              <tr>
                <td height=20><A href="send_zn_message.asp" target="main">����վ����Ϣ</A></td>
              </tr>
                <tr><td height=20><A href="zn_message.asp" target="main">����վ����Ϣ</A></td>
                </tr>
				<tr>
                  <td height=20><A href="phone_message.asp" target="main">�����ֻ�����</A></td>
                </tr>
				<%end if
				if AdminPurview=1 or AdminPurview_class3=1 then%>
<tr><td height=20><A href="Admin_FriendLinks.asp" target="main">��������</A></td>
</tr>
<%end if%>
<tr><td height=20><A href="qqadmin.asp" target="main">QQ�ͷ�����</A></td>
</tr>
<%if AdminPurview=1 or AdminPurview_class4=1 then%>
<tr>
  <td height=20><A href="heiguan.asp" target="main">�ٷ�����������</A></td>
</tr>
<tr>
  <td height=20><A href="hei.asp" target="main">����������</A></td>
</tr>
<%end if
if AdminPurview=1 or AdminPurview_class5=1 then%>
<tr>
  <td height=20><A href="viewreturn.asp" target="main">���Թ���</A></td>
</tr>
<%end if
if AdminPurview=1 or AdminPurview_class7=1 then%>
<tr>
  <td height=20><A href="exesql.asp" target="main">�Զ���ִ��SQL</A></td>
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
          <span>���¹���</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu1'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height="20"><a href=Admin_ArticleAdd1.asp target=main>�������</a> | <a href=Admin_ArticleManage.asp target=main>���¹���</a></td>
</tr>



<tr>
  <td height=20><a href=Admin_ArticleRecyclebin.asp target=main>����վ����</a></td>
</tr>

<tr><td height=20><a href=Admin_Class_Article.asp?Action=Add target=main>������Ŀ���</a> | <a href=Admin_Class_Article.asp target=main>����</a></td></tr>

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
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_8.gif" id=menuTitle8 onClick="showsubmenu(422)" style="cursor:hand;"><span>��ֵ������</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu422'><div class=sec_menu style="width:158">
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr>
            <td height=20><A href="Admin_Jicar.asp" target="main">��ֵ������</A></td>
          </tr>
          <tr>
            <td height=20><a href="Admin_Okjicar.asp" target="main">�㿨��ֵ��¼</a></td>
          </tr>
          
           <tr>
            <td height=20><A href="Admin_Jiagesear.asp" target="main">�㿨�����ѯ</A> </td>
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
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_8.gif" id=menuTitle8 onClick="showsubmenu(203)" style="cursor:hand;"><span>��������</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu203'><div class=sec_menu style="width:158">
        <table cellpadding=0 cellspacing=0 align=center width=130>
		   <tr>
            <td height=20><a href=admin_delrecord.asp target=main >һ��������־</a></td>
          </tr>
		   <tr>
            <td height=20><a href=admin_deltaobao.asp target=main >һ����������</a></td>
          </tr>
          <tr>
            <td height=20><a href=admin_Mission.asp target=main>��������</a></td>
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
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_8.gif" id=menuTitle8 onClick="showsubmenu(276)" style="cursor:hand;"><span>��ֵ�����������</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu276'><div class=sec_menu style="width:158">
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr>
            <td height=20><a href=admin_MyMission4.asp target=main>��ֵ�����������</a></td>
          </tr>
		   <tr>
            <td height=20><a href=admin_MyMissionjie4.asp target=main>�ѽӳ�ֵ��������</a></td>
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
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_8.gif" id=menuTitle8 onClick="showsubmenu(2)" style="cursor:hand;"><span>�Ա����������</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu2'><div class=sec_menu style="width:158">
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr>
            <td height=20><a href=admin_MyMission.asp target=main>�Ա��������</a></td>
          </tr>
		   <tr>
            <td height=20><a href=admin_MyMissionjie.asp target=main>�ѽ��Ա��������</a></td>
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
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_8.gif" id=menuTitle8 onClick="showsubmenu(236)" style="cursor:hand;"><span>�ղ��������</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu236'><div class=sec_menu style="width:158">
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr>
            <td height=20><a href=admin_MyMission5.asp target=main>�ղ��������</a></td>
          </tr>
		   <tr>
            <td height=20><a href=admin_MyMissionjie5.asp target=main>�ѽ��ղ�����</a></td>
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
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_8.gif" id=menuTitle8 onClick="showsubmenu(24)" style="cursor:hand;"><span>�������������</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu24'><div class=sec_menu style="width:158">
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr>
            <td height=20><a href=admin_MyMission1.asp target=main>�����������</a></td>
          </tr>
		    <tr>
            <td height=20><a href=admin_MyMissionjie3.asp target=main>�ѽ������������</a></td>
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
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_8.gif" id=menuTitle8 onClick="showsubmenu(42)" style="cursor:hand;"><span>�û�����</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu42'><div class=sec_menu style="width:158">
        <table cellpadding=0 cellspacing=0 align=center width=130>
		<%if AdminPurview=1 or AdminPurview_class16=1 then%>

          <tr>
            <td height=20><a href=usermanage.asp target=main>�û�����</a></td>
          </tr>
           <%end if%>
          
           <tr>
             <td height=20><a href=chengfaman.asp target=main>���͹���</a></td>
           </tr>
		   <%if AdminPurview=1 or AdminPurview_class17=1 then%>

           <tr>
             <td height=20><a href=admin_recordm.asp target=main>���������¼</a></td>
           </tr>
           <%end if
		   if AdminPurview=1 or AdminPurview_class18=1 then%>

           <tr>
            <td height=20><a href=admin_record.asp target=main>��Ա������¼</a></td>
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
          <span>������� </span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu3'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>
               <tr>
			    <td height=20><a href="jieducm_car.asp" target=main>��վ�㿨����</a></td>
		      </tr>
			   <tr>
			     <td height=20><a href="jieducm_vipcar.asp" target=main>VIP������</a></td>
		      </tr>
			                 <tr>
                 <td height=20><a href="jieducm_pay.asp" target="main">�����㽻���г�</a></td>
               </tr>
			   <tr>
                 <td height=20><a href="AutoCZ_Manage.asp" target=main><font color="red">ת�˳�ֵ����</font></a></td>
               </tr>
		      <tr>
			    <td height=20><a href="admin_bankroll.asp" target=main>���ּ�¼</a></td>
		      </tr>
               <tr>
                 <td height=20><a href="chongrecord.asp" target=main>��ֵ��¼</a></td>
               </tr>
               <tr>
			    <td height=20><a href="jieducm_jiesuan.asp" target=main>ƽ̨������</a></td>
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
          <span>������� </span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu323'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>
              
			  <tr>
			    <td height=20><a href="tribemanage.asp" target=main>�������</a></td>
		      </tr>
               <tr>
			    <td height=20><a href="tribebook.asp" target=main>�������Թ���</a></td>
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
          <span>����Ա����</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu4'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>

              <tr>
                <td height=20><a href=Admin_Admin.asp?Action=Add target=main>����Ա���</a> 
                  | <a href=Admin_Admin.asp target=main>����</a></td>
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
          <span>������</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu6'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr>
                <td height=20><a href="HL_admin_gg.asp" target="main">������</a></td>
          </tr>
<tr>
                <td height=20><a href="HL_admin_gg.asp?action=tj" target="main">��ӹ��</a></td>
          </tr>
          <tr>
                <td height=20><a href="HL_admin_manwei.asp" target="main">���λ����</a></td>
          </tr>
          <tr>
                <td height=20><a href="HL_admin_ggmanlei.asp" target="main">���������</a></td>
          </tr>
                    <tr>
                <td height=20><a href="guanggao.asp" target="main">��ҳFLASH���</a></td>
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
          <span>���߹���</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu62'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr>
                <td height=20><a href="shenshu.asp" target="main">���߹���</a></td>
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
          <span>SQLע�����</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu622'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr>
                <td height=20><a href="sqlInConfig.asp?moduleid=320" target="main">ϵͳ����</a></td>
          </tr>
		  <tr>
                <td height=20><a href="sqlInList.asp?moduleid=321" target="main">�鿴��Ϣ</a></td>
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
          <span>��վ��Ϣ</span></td>
  </tr>
  <tr>
    <td>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20 valign="top"><br>
���򿪷���<a href="http://www.688880.com" target="_blank">��������</a><br>
����QQ��1662424999<br>
&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
</table>
	  </div>
	</td>
  </tr>
</table>
</body>
</html>