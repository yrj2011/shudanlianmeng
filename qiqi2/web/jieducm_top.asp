<%
Url=split(request.servervariables("script_name"),"/")
FileName=Url(ubound(Url))
call alipayto() 
File=Left(FileName,InstrRev(FileName,".")-1)
%>
 
<DIV style="MARGIN-TOP: 0px" id=sitenav>
<DIV style="PADDING-BOTTOM: 0px; PADDING-LEFT: 0px; PADDING-RIGHT: 0px; PADDING-TOP: 0px" class="wrapper bold">
<DIV id=headnt class=info><A style="CURSOR: pointer" onclick='$("#sitenav").slideToggle("slow")'></A><MARQUEE onmouseover=this.stop()  onmouseout=this.start() scrollAmount=2 direction=left width="800" height=28><font color=#ff5e10><%=stupjieducm_gonggao%></font></MARQUEE></DIV></DIV></DIV>
<table width="100%" border="0" cellpadding="0" cellspacing="0" background="images/index_top_bg.jpg">
  <tr>
    <td><TABLE width=969 border=0 align=center cellPadding=2 cellSpacing=0  >
  <TR>
    <TD width=433><A href="<%=stupurl%>" ><IMG alt=<%=stupname%> src="<%=stuplogo%>" border=0></A> 
	</TD>
    <TD vAlign=top width=528>
      <TABLE cellSpacing=0 cellPadding=0 width=535 align=right border=0>
        <TR>
          <TD height=32 align="right"><span class="yell_font">��ӭ����</span>
            <%if session("useridname")="" then%>
<A href="login.asp" target="_top">��½</A> |<A  href="register.asp" target="_top"> ע��</A> |
<%else
Sql = "select fabudian,cunkuan,jifei,vip,tribeid from "&jieducm&"_register where username='"&session("useridname")&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Rs.Eof Then
	response.Redirect("login.asp")				
Else
    fabudian=rs("fabudian")
	cunkuan=rs("cunkuan")
	jifei=rs("jifei")
	vip=Replace(Trim(rs("vip")),"'","''")
	tribeid=rs("tribeid")
end if
rs.close	
%>
 <A href="user/" target="_top"><SPAN><%=session("useridname")%></SPAN> </A>
 <%if vip="1" then %><img src="images/jieducm_vip.gif"  alt="���VIP,����ֵ��<%call vipxinyu(session("useridname"))%>"  border="0"/><% end if%> |  <A onClick="return confirm('ȷ���˳�������');"  href="user/exit.asp" target="_top">[�˳�]</A>  | 
 <% end if%>
 <A href="user/" target="_top"> <img src="images/jieducm_member.png" border="0" />��Ա���� </A>| <A href="#" title=��ӵ��ղؼ� onclick="window.external.addFavorite('<%=stupurl%>','<%=stupname%>')" ><img src="images/jieducm_cang.png" width="23" height="23"  border="0"/>�ղر�վ</A> | <A href="help/link.asp" target="_top"><img src="images/jieducm_asd.png" width="22" height="23"  border="0"/>������</A></TD>
        </TR>
        <TR>
          <TD height=20>
		  <DIV id=index_top_r2>     
      <LI style="BACKGROUND: url(images/jieducm_top_btn6.jpg) no-repeat"><A onClick="return Button3_onclick()" href="javascript:void(0)">�ͷ�����</A> 
	  <LI style="BACKGROUND: url(images/jieducm_top_btn5.jpg) no-repeat"><A  href="user/ListTop.asp">ˢ������</A> </LI>
	  <LI style="BACKGROUND: url(images/jieducm_top_btn4.jpg) no-repeat"><A href="user/sms.asp">�ֻ�����</A> </LI>
	  <LI style="BACKGROUND: url(images/jieducm_top_btn3.jpg) no-repeat"><A href="user/MoneyOrPush.asp">��Ҫ�һ�</A> </LI>
	  <LI style="BACKGROUND: url(images/jieducm_top_btn2.jpg) no-repeat"><A  href="user/mai.asp">�򷢲���</A> </LI>
	  <LI style="BACKGROUND: url(images/jieducm_top_btn1.jpg) no-repeat"><A href="user/md5_pay.asp">�˺ų�ֵ</A> </LI>
      </LI></DIV>
		  </TD>
          </TR></TABLE></TD></TR></TABLE></td>
  </tr>
</table>

<DIV class=index_menu style="MARGIN: 0px auto">
      <DIV id=list>
      <UL>
        <LI class=white><A <%if file<>"new" and file<>"news" and file<>"newmore" then%>class="li1 white"<%else%>class="li0 " <% end if%>  href="index.asp" id="a0"  onmouseover="Mea(0);" onMouseOut="setAuto()"><%if file<>"new" and file<>"news" and file<>"newmore" then%><SPAN  style="COLOR: #000000">��&nbsp;ҳ</SPAN><%else%>��&nbsp;ҳ<%end if%></A> </LI>
        <LI class=white><A class="li0 " href="taobao/" id="a2" onmouseover="Mea(2);" onMouseOut="setAuto(2)">�Ա�����</A> </LI>
        <LI class=white><A class="li0 " href="pai/" id="a4" onmouseover="Mea(4);" onMouseOut="setAuto(4)">��������</A> </LI>
		<LI class=white><A class="li0 " href="../user/soft.asp" id="a3" onmouseover="Mea(3);" onMouseOut="setAuto(3)">�������</A> </LI>
        <LI class=white><A class="li0 " href="../user/tuoguan.asp" id="a5" onmouseover="Mea(5);" onMouseOut="setAuto(5)">�����й�</A></LI>
        <LI class=white><A class="li0 " href="chongzhi/" id="a6" onmouseover="Mea(6);" onMouseOut="setAuto(6)">��ֵ������</A> </LI>
        <LI class=white><A class="li0 "  href="union/" id="a7" onmouseover="Mea(7);" onMouseOut="setAuto(7)">�ƹ�׬Ǯ</A> </LI>
        <LI class=white><A class="li0 " href="tribe/" id="a8" onmouseover="Mea(8);" onMouseOut="setAuto(8)">���佻��</A> </LI>
		<LI class=white><A class="li0 " href="teach.asp " id="a1" onmouseover="Mea(1);" onMouseOut="setAuto(1)">�����̳�</A> </LI>
        <LI class=white><A <%if file="new" or file="news" or file="newmore" then%> class="li1 white"<%else%>class="li0" <% end if%>  href="new.asp" id="a9" onmouseover="Mea(9);" onMouseOut="setAuto()"><%if file="new" or file="news" or file="newmore" then%> <SPAN  style="COLOR: #000000">�������� </SPAN><%else%>��������<% end if%></A> </LI>
		</UL></DIV></DIV>
		
			
		<style>
#sideBuoy { position: fixed; _position: absolute; z-index: 9; top: 205px; _top: expression(205+documentElement.scrollTop +"px"); left: 50%; width: 81px; height: 173px; margin-left: 500px; background: #82cbec url(/images/sideBuoy-qq.jpg) no-repeat 0 30px;}
.sideBuoy-jy { display: block; width: 100%; height: 30px; text-align: center; line-height: 30px;}
.sideBuoy-qq { display: block; width: 100%; height: 112px;}
.sideBuoy-top { display: block; width: 100%; height: 31px; text-align: center; line-height: 30px;}
.sideBuoy-detail { display: none; position: absolute; left: -156px; top: 0; overflow: hidden; width: 156px; height: 173px; background: #ffeeda; text-align: center;}
.sideBuoy-detail .qq-list { padding: 8px 0 10px;}
.sideBuoy-detail .qq { height: 22px; padding: 5px 0; line-height: 22px;}
.sideBuoy-detail .qq img { vertical-align: middle;}
</style>
<div id="sideBuoy">
    <a class="sideBuoy-jy" href="/help/viewreturn.asp" tppabs="/help/viewreturn.asp" target="_blank">���н���</a>
    <a class="sideBuoy-qq" href="#"></a>
    <a class="sideBuoy-top" href="#">���ض���</a>
    <div class="sideBuoy-detail">
        <div class="qq-list" ID="qqAllList"><div class="qq">����������ѯ</div>
<div class="qq">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" valign="middle"> <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=1662424999&site=qq&menu=yes"><img src="http://wpa.qq.com/pa?p=1:1662424999:41"  border="0"  /><BR />1662424999</A></td>
  </tr>
</table>
</div>
</div>
<br/><br/><br/><br/><br/><br/><div class="service-time">����ʱ�䣺9:00~21:00<br><br/>
�ǹ���ʱ�������������
</div>
    </div>
</div>
<script>
var flag,
    timer;
$('.sideBuoy-qq').hover(function(){
    flag = false;
    //createqq();
    $('.sideBuoy-detail').show();
},function(){
    flag = true;
    timer = setTimeout(function(){
        if(flag){
            $('.sideBuoy-detail').hide();
        }
    },200) 
})
$('.sideBuoy-detail').hover(function(){
    flag = false;
    $(this).show();
},function(){
    var $self = $(this);
    flag = true;
    timer = setTimeout(function(){
        if(flag){
            $self.hide();
        }
    },200) 
})
</script>