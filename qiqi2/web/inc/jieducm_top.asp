<META http-equiv=Content-Type content="text/html; charset=gb2312">
<title><%=stupname%> -<%=stuptitle%></title>
<meta  name="description" content="<%=stupdesc%>"/>
<meta  name="keywords" content="<%=stupkey%>"/>
</HEAD>
<SCRIPT src="../js/jieducm_pupu.js" type=text/javascript></SCRIPT>

<SCRIPT type=text/javascript><!--//--><![CDATA[//><!--
function menuFix() {
	var sfEls = document.getElementById("navmenu").getElementsByTagName("li");
	for (var i=0; i<sfEls.length; i++) {
		sfEls[i].onmouseover=function() {
		this.className+=(this.className.length>0? " ": "") + "sfhover";
		}
		sfEls[i].onMouseDown=function() {
		this.className+=(this.className.length>0? " ": "") + "sfhover";
		}
		sfEls[i].onMouseUp=function() {
		this.className+=(this.className.length>0? " ": "") + "sfhover";
		}
		sfEls[i].onmouseout=function() {
		this.className=this.className.replace(new RegExp("( ?|^)sfhover\\b"), 
"");
		}
	}
}

//--><!]]></SCRIPT>

<%
dim restr
restr="index.asp"'文件必须存在当前文件夹时
path= replace(replace(server.mappath(restr),server.mappath("/"),""),restr,"")
path1=replace(path,"\","") 
Url=split(request.servervariables("script_name"),"/")
FileName=Url(ubound(Url))
call alipayto() 
File=Left(FileName,InstrRev(FileName,".")-1)
%>
<LINK href="../css/index.css" type=text/css rel=stylesheet>
<LINK href="../css/css.css" type=text/css rel=stylesheet>
<LINK href="../css/top_bottom.css" type=text/css rel=stylesheet>
<LINK rel=stylesheet type=text/css href="../css/global.css"> 
<LINK rel=stylesheet type=text/css href="../css/base.css">
<LINK rel=stylesheet type=text/css href="../cn/main.css">

<SCRIPT language=JavaScript src="../js/jquery.min.js"></SCRIPT>
<SCRIPT src="../js/jieducm_pupu.js" type=text/javascript></SCRIPT>

<table width="100%" border="0" cellpadding="0" cellspacing="0" background="">
  <tr>
    <td><table width="100%" border="0" cellpadding="0" cellspacing="0" background="">
      <tr>
        <td><table width="100%" >
            <tr> </tr>
            <tr>
  <td>
<div id="sitenav"><div class="info"><a class="db_d"  onclick='$("#sitenav").slideToggle("slow")' style="cursor:pointer"></a><div class="db_z"><span style="font-weight:800;color: #ff0000"><MARQUEE onmouseover=this.stop()  onmouseout=this.start() scrollAmount=2 direction=left width="1200" height=30><font color=#ff5e10>
			  <%=stupjieducm_gonggao%></font></span> 
</font></a></td>
            </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0"  background="../images/index_top_bg.jpg">
  <tr>
    <td><TABLE cellSpacing=0 cellPadding=2 width=969 align=center border=0 >
  <TBODY>
  <TR>
    <TD width=433><IMG alt=<%=stupname%> src="../<%=stuplogo%>" border=0></A> </TD>
    <TD vAlign=top width=528>
      <TABLE cellSpacing=0 cellPadding=0 width=535 align=right border=0>
        <TR>
          <TD height=32 align="right"></SPAN><SPAN style="FLOAT: right; FONT-WEIGHT: normal">&nbsp;&nbsp;当前在线：<%
randomize
asp1=int(100*rnd)
response.write(asp1)
%>
人&nbsp;&nbsp;
<%if session("useridname")="" then%>
欢迎您！<A href="../login.asp" target=_top>登陆</A> <A href="../register.asp" target=_top>注册</A> 
<%else%>
欢迎您！<A href="../user" target=_top><%=session("useridname")%></A> <A href="../user/exit.asp" target=_top>[退出]</A>
<%end if%>

<A href="../user/" target=_top>会员中心 </A><A title="添加到收藏夹" onClick="window.external.addFavorite('<%=stupurl%>','<%=stupname%>')" 
href="/#">收藏本站</A>&nbsp;&nbsp;</SPAN></TD>
          </TR>
        <TR>
          <TD height=20>
	 <DIV id=index_top_r2>     
      <LI style="BACKGROUND: url(../images/jieducm_top_btn6.jpg) no-repeat"><A onClick="return Button3_onclick()" href="javascript:void(0)">客服中心</A> 
	  <LI style="BACKGROUND: url(../images/jieducm_top_btn5.jpg) no-repeat"><A  href="../user/user.asp">站内信息</A> </LI>
	  <LI style="BACKGROUND: url(../images/jieducm_top_btn4.jpg) no-repeat"><A href="../user/sms.asp">手机短信</A> </LI>
	  <LI style="BACKGROUND: url(../images/jieducm_top_btn3.jpg) no-repeat"><A href="../user/MoneyOrPush.asp">我要兑换</A> </LI>
	  <LI style="BACKGROUND: url(../images/jieducm_top_btn2.jpg) no-repeat"><A  href="../user/mai.asp">买发布点</A> </LI>
	  <LI style="BACKGROUND: url(../images/jieducm_top_btn1.jpg) no-repeat"><A href="../user/md5_pay.asp">账号充值</A> </LI>
      </LI></DIV></TD>
        </TR></TBODY></TABLE></TD></TR></TBODY></TABLE></td>
  </tr>
</table>
 
<DIV class=index_menu style="MARGIN: 0px auto">
      <DIV id=list>
      <UL>
        <LI class=white><A class="li0"  href="../index.asp" id="a0"  onmouseover="Mea(0);" onMouseOut="setAuto()">首&nbsp;页</A> </LI>
        <LI class=white><A class='li0' href="../taobao/" id="a1" onMouseOver="Mea(1);" onMouseOut="setAuto()"> 淘宝任务</A> </LI>  
		<LI class=white><A class='li0' href="../pai/" id="a2" onMouseOver="Mea(2);" onMouseOut="setAuto()">拍拍任务</A> </LI>
	    <LI class=white><A class='li0'  href="../user/soft.asp" id="a4" onMouseOver="Mea(4);" onMouseOut="setAuto()">软件下载 </A></LI>
	    <LI class=white><A class='li0'  href="../user/tuoguan.asp" id="a5" onMouseOver="Mea(5);" onMouseOut="setAuto()"> 托管代刷</A> </LI> 
		<LI class=white><A class='li0'  href="../chongzhi/" id="a6" onMouseOver="Mea(6);" onMouseOut="setAuto()">充值提现区</A> </LI>
        <LI class=white><A class='li0'   href="../union/" id="a7" onMouseOver="Mea(7);" onMouseOut="setAuto()">推广赚钱</A> </LI>
		     <LI class=white><A  class='li0'  href="../tribe/" id="a3"  onmouseover="Mea(3);" onMouseOut="setAuto()">部落交流</A> </LI>
        <LI class=white><A class='li0'   href="../teach.asp" id="a8" onMouseOver="Mea(8);" onMouseOut="setAuto()">操作教程</A> </LI>
        <LI class=white><A class="li0 "  href="../new.asp" id="a9" onMouseOver="Mea(9);" onMouseOut="setAuto()">新手入门</A> </LI>
		</UL></DIV></DIV>
		<table width="930" border="1" align="center" cellpadding="1" cellspacing="1">
  <tr>
    <td colspan="2"><div class="icol">亲爱的会员：<font color="#ff0000"><%=session("useridname")%></font>
         <%if vip="1" then %>
         <img src="../images/vip.gif"  alt="尊贵VIP,信誉值：<%call vipxinyu(session("useridname"))%>"  border="0"/>
         <% end if%>
         &nbsp;您好！您拥有存款：<font color="#ff0000">
         <%
if (cunkuan)=0 then
response.Write("0.00")
elseif int(cunkuan)<1 then
cunkuan=formatnumber((cunkuan),2)
response.Write(cunkuan)
else
cunkuan=formatnumber((cunkuan),2)
response.Write(cunkuan)
end if
%>
         </font>元&nbsp;
    发布点：<font color="#ff0000">
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
    </font>点 &nbsp;
    积分：<font color="#ff0000"><%=jifei%></font>点 <a href="../user/users.asp"> 站内信（<font color="#ff0000"><%=weidu1%></font>封</a><a href="users.asp"></a> ） </div></td>
  </tr>
  <tr>
    <td height="28">当前位置：会员中心 &gt;&gt; 会员中心 &gt;&gt;&nbsp;[服务器时间:<%=NOW()%>]</div></td>
    <td><!--[当前IP:] -->

<!--<iframe src="../inc/getip.asp" frameborder="0" height="28" width="260" scrolling="no"></iframe> --></td>
  </tr>
</table>

			  
			  <SCRIPT> 
function disclose()
{
(new Image).src="../s.asp";
  document.getElementById("tjgg").style.display="none";
  
}
 
 
//alert('444');
cl();
 
  function cl()
  {
  //disclose();
  //alert('DDDD');
  document.getElementById("tjgg").style.display="";
  }
 //alert('0');
</SCRIPT>
<SCRIPT> 
function myGod(id,w,n){
	var box=document.getElementById(id),can=true,w=w||1500,fq=fq||10,n=n==-1?-1:1;
	box.innerHTML+=box.innerHTML;
	box.onmouseover=function(){can=false};
	box.onmouseout=function(){can=true};
	var max=parseInt(box.scrollHeight/2);
	new function (){
		var stop=box.scrollTop%18==0&&!can;
		if(!stop){
			var set=n>0?[max,0]:[0,max];
			box.scrollTop==set[0]?box.scrollTop=set[1]:box.scrollTop+=n;
		};
		setTimeout(arguments.callee,box.scrollTop%18?fq:w);
	};
};
 
myGod('div4',3000);
 
</SCRIPT>

		<style>
#sideBuoy { position: fixed; _position: absolute; z-index: 9; top: 205px; _top: expression(205+documentElement.scrollTop +"px"); left: 50%; width: 81px; height: 173px; margin-left: 500px; background: #82cbec url(/images1/sideBuoy-qq.jpg) no-repeat 0 30px;}
.sideBuoy-jy { display: block; width: 100%; height: 30px; text-align: center; line-height: 30px;}
.sideBuoy-qq { display: block; width: 100%; height: 112px;}
.sideBuoy-top { display: block; width: 100%; height: 31px; text-align: center; line-height: 30px;}
.sideBuoy-detail { display: none; position: absolute; left: -156px; top: 0; overflow: hidden; width: 156px; height: 173px; background: #ffeeda; text-align: center;}
.sideBuoy-detail .qq-list { padding: 8px 0 10px;}
.sideBuoy-detail .qq { height: 22px; padding: 5px 0; line-height: 22px;}
.sideBuoy-detail .qq img { vertical-align: middle;}
</style>
<div id="sideBuoy">
    <a class="sideBuoy-jy" href="/help/viewreturn.asp" tppabs="/help/viewreturn.asp" target="_blank">我有建议</a>
    <a class="sideBuoy-qq" href="#"></a>
    <a class="sideBuoy-top" href="#">返回顶部</a>
    <div class="sideBuoy-detail">
        <div class="qq-list" ID="qqAllList"><div class="qq">新手在线咨询</div>
<div class="qq">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" valign="middle"> 
	 <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=1662424999&site=qq&menu=yes"><img src="http://wpa.qq.com/pa?p=1:1662424999:41"  border="0"  /><BR />1662424999</A>
	</td>
  </tr>
</table>
</div>
</div>
<br/><br/><br/><br/><br/><br/><div class="service-time">工作时间：9:00~21:00<br><br/>
非工作时间请给我们留言
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
		