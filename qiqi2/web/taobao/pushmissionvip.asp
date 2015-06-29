<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->

<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：捷度传媒
'官方主页：http://www.jieducm.cn
'技术支持：robin 
'QQ;616591415
'电话：13598006137
'版权：版权属于捷度传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2008－2011 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
if vip<>1 then
	Response.Write("<script>alert('你还不是vip会员!不能发布vip任务！');window.location.href='../user/login.asp';</script>")
    response.End()
elseif session("czm")="" then
 response.Redirect("code.asp?id=fa")
 response.End()
end if
call myww(1)
response.charset="gb2312"
action=request.QueryString("go")
if action="check" then
url = request("url")
nickname = trim(request("nickname"))
content = GetNickName(url)
if nickname<>content then
echo "no"
else
echo "yes"
end if
response.End()
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh" lang="zh" dir="ltr">
<head>
<title><%=stupname%> -<%=stuptitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="../css/style.css" type=text/css rel=stylesheet>
<LINK href="../css/base.css" type=text/css rel=stylesheet>
<LINK href="../css/layout.css" type=text/css rel=stylesheet>
 <script type="text/javascript" src="../js/common5.js"></script>
 <script type="text/javascript" src="../js/jieducm_jsdate.js"></script>
</head>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=jieducm_toptao.asp-->
<div align="center">  
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 800px">
<UL id=task_nav>
 <li><a  href="index.asp">淘宝互动区</a> </li>
    <li><a style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="pushmissionvip.asp">发布任务</a> </li>
  <LI><A href="ReMission.asp">已接任务</A> </LI>
  <LI><A href="MyMission.asp">已发任务</A> </LI>
  <LI><A  href="MyWw.asp">绑定店铺</A> </LI>
  <LI><A href="Buyhao.asp">绑定买号</A> </LI>
  <LI><A href="MySave.asp">任务仓库</A> </LI>
   <li><a href="../user/Express.asp">生成快递单</a> </li>
  </UL></DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG src="../images/task_nav_line.gif"></DIV>
</DIV>
</div>
  <table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>
        <table width="910"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            
            <td valign="top">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr height="1">
                  <td height="5"></td>
      </tr>
  </table>
              <table width="910"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            
            <td valign="top">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr height="1">
                  <td height="5"></td>
      </tr>
  </table>
              <DIV class=page>
<DIV class=addtask-wrap>
<DIV class=inner>
发布任务前须知： 
<UL>
  <LI>1. 发布一个任务会扣除您的存款，请保证您的平台存款充足； 
  <LI>2. 请保证您在本站有足够的存款，您所发布任务价格金额从存款里面扣除，有人接到任务后做作为商品款返回给你。 </LI>
    
  <li>3. (任务价格是包含邮费的总价格)1-40元，扣一个发布点；41-100元扣两个发布点；101-200元，扣三个发布点；201-500元，扣四个发布点；501-1000元扣五个发布点</li>
  </UL></DIV></DIV><br>
<DIV class=addtask-wrap>
<DIV class=inner>
  <TABLE cellSpacing=0 cellPadding=0 border=0>
  <TBODY>
     <FORM name=form id="form" method=post action="jieducm_faokvip.asp" >
           <input name="fa" type="hidden" value="ok">
  <TR>
    <TH><div align="right">请先选择连发几个任务：</div></TH>
    <TD><LABEL><INPUT id="num" onclick=javascript:gouwuce(1); type=radio CHECKED value="1" name=num>
                                    一链刷</LABEL>
                                     <LABEL><INPUT id="num"  onclick=javascript:gouwuce(2); type=radio  value="2" name=num>
                                    二链刷</LABEL>
                                    <LABEL><INPUT id="num" onclick=javascript:gouwuce(3); type=radio  value="3" name=num>
                                    三链刷</LABEL>
                                   <LABEL > <INPUT id="num" onclick=javascript:gouwuce(4); type=radio  value="4" name=num>
                                    四链刷</LABEL>
                                    <LABEL><INPUT id="num" onclick=javascript:gouwuce(5); type=radio  value="5" name=num>
                                    五链刷</LABEL></TD></TR>
  <TR>
    <TH><div align="right">掌&nbsp; 柜&nbsp; 名：</div></TH>
    <TD ><%call shopname(1)%></TD>
  </TR>
  
  <TR>
    <TH><div align="right">任务1：</div></TH>
    <TD >商品任务价：<input name=Price1 id=Price1 size="10" maxlength=6  onKeyUp="if(isNaN(value))execCommand('undo')">
	商品地址：  <input id=ProUrl1 maxlength=255 name=ProUrl1 onBlur="check(this)">	</TD></TR>
  <TR id=renwu2 style="DISPLAY: none">
    <TH><div align="right">任务2：</div></TH>
    <TD >商品任务价：<input name=Price2 id=Price2 size="10" maxlength=6 onKeyUp="if(isNaN(value))execCommand('undo')">
	商品地址：  <input id=ProUrl2 maxlength=255 name=ProUrl2 onBlur="check(this)" >	</TD></TR>
	<TR id=renwu3 style="DISPLAY: none">
    <TH><div align="right">任务3：</div></TH>
    <TD >商品任务价：<input name=Price3 id=Price3 size="10" maxlength=6 onKeyUp="if(isNaN(value))execCommand('undo')">
	商品地址：  <input id=ProUrl3 maxlength=255 name=ProUrl3  onBlur="check(this)">	</TD></TR>
	<TR id=renwu4 style="DISPLAY: none">
    <TH><div align="right">任务4：</div></TH>
    <TD >商品任务价：<input name=Price4 id=Price4 size="10" maxlength=6 onKeyUp="if(isNaN(value))execCommand('undo')">
	商品地址：  <input id=ProUrl4 maxlength=255 name=ProUrl4  onBlur="check(this)">	</TD></TR>
	<TR id=renwu5 style="DISPLAY: none">
    <TH><div align="right">任务5：</div></TH>
    <TD >商品任务价：<input name=Price5 id=Price5 size="10" maxlength=6 onKeyUp="if(isNaN(value))execCommand('undo')">
	商品地址：  <input id=ProUrl5 maxlength=255 name=ProUrl5  onBlur="check(this)">	</TD></TR>
 		<TR id=renwu6 style="DISPLAY: none">
    <TH><div align="right">任务6：</div></TH>
    <TD >商品任务价：<input name=Price6 id=Price6 size="10" maxlength=6 onKeyUp="if(isNaN(value))execCommand('undo')">
	商品地址：  <input id=ProUrl6 maxlength=255 name=ProUrl6  onBlur="check(this)">	</TD></TR>
         <TR>
    <TH><div align="right">是否来路搜索任务：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><label>
      <input type="checkbox" name="jieducm_IfComeSet" id="jieducm_IfComeSet" value="1" onClick="javascript:SwitchDivDisplay('showinfo_c1')" >
      来路搜索任务
    </label>
    (选择此项额外扣1个发布点)</TD>
  </TR>
      <tr id=showinfo_c1 style="DISPLAY: none">
      <TH><div align="right">来路任务选择：</div></TH>
      <TD colspan="3" 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><label><INPUT 
      onclick=setTaskid(1) value=1 CHECKED type=radio name=IfComeSet>淘宝店铺搜索</label><SPAN 
      style="MARGIN-LEFT: 20px"> 是在淘宝搜索店铺进入您店里的流量 </SPAN></SPAN><BR><SPAN 
      style="MARGIN-RIGHT: 10px"><label><INPUT onclick=setTaskid(2) value=2 type=radio 
      name=IfComeSet>淘宝宝贝搜索</label><SPAN style="MARGIN-LEFT: 20px"> 是在淘宝直接搜索宝贝进入的 </SPAN></SPAN><BR></TD>
    </TR>
	 <tr id=showinfo_c2 style="DISPLAY: none">
      <TH><div align="right"><SPAN id=tt1>搜索店铺关键字</SPAN>：</div></TH>
      <TD colspan="3" 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
	<INPUT style="WIDTH: 290px" id=ComeKey class=fa-input maxLength=50  type=text name=ComeKey><SPAN id=ttt1 
      class=fa-righttext>此处输入您需要告诉接手人商品在自己店铺内的什么地方可以找到，例如，进入店铺内搜索“100元面子卡”商品，第一个就是。 
      </SPAN></TD>
    </TR>
    <tr id=showinfo_c3 style="DISPLAY: none">
      <TH><div align="right"><SPAN id=tt2>搜索提示</SPAN>：</div></TH>
      <TD colspan="3" 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"> 
	<INPUT style="WIDTH: 290px" id=ComeNote class=fa-input maxLength=50  type=text name=ComeNote><SPAN id=ttt2 
      class=fa-righttext>此处填写提示信息，说明店铺在关键字搜索结果列表中的位置，商品在店铺首页中大概位置等信息，比如：店铺在结果列表第二个，商品在商品首页第二排中间一个等等。</SPAN></TD>
    </TR>
		
		<tr>
          <TH><div align="right">增加发布点：</div></TH>
          <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
            <label>
              <select name="addfabu">
                <option value="0">不增加</option>
                <option value="1">增加1个</option>
                <option value="2">增加2个</option>
                <option value="3">增加3个</option>
                <option value="4">增加4个</option>
                <option value="5">增加5个</option>
              </select>
              </label>
          
         &nbsp; (套餐任务必选)</TD>
        </TR>
        <TH><div align="right">价格是否相等：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><input name="isprice" type="radio" id="isprice" value="金额相等" checked>
                              <span class="font12l"> 金额相等</span> <input type="radio" name="isprice" id="isprice" value="需修改价格">
                               <span class="font12l">需修改价格</span>&nbsp; 任务价格和包括运费的商品总价格是否相等！</TD>
  </TR>
  <TR>
    <TH><div align="right">动态评分：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><input name="play" type="radio" value="全部打5分" checked>
                              <span class="font12l">全部打5分</span> <input type="radio" name="play" value="全部不打分">
                             <span class="font12l"> 全部不打分</span></TD>
  </TR>
   <TR>
    <TH></TH>
    <TD><FONT color=#ff0000><STRONG>* 
      发布延时收货的任务，平台免费提供物流单号，并强制买家延时收货</STRONG></FONT></TD></TR>
	  <TR>
    <TH><div align="right">好评方式：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
      <label>
        <input name="ping" type="radio" value="带字好评" checked onclick=javascript:jieducm_ping(1);> 带字好评</label>
      <label>
        <input type="radio" name="ping" value="不带字好评" onclick=javascript:jieducm_ping(2);>不带字好评 </label>
		    <label>
        <input type="radio" name="ping" value="自定义评语" onclick=javascript:jieducm_ping(3);>自定义评语 </label>
			    <label>
        <input type="radio" name="ping" value="默认好评" onclick=javascript:jieducm_ping(4);>默认好评(不评价)</label>    </TD>
  </TR>
  <TR id=pinghistory style="DISPLAY: none">
    <TH><div align="right">自定义评语：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
        <INPUT   name=pingtxt class=txtinput id=pingtxt style="WIDTH: 290px" maxlength="100">
    <select name="lsx"  style="width:150px;" onChange="jieducm_pingtxt(this.options[this.options.selectedIndex].value);"> 
								<option value="">请选择历史自定义评语</option>
<%Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select top 10 pingtxt from "&jieducm&"_zhongxin where username='"&session("useridname")&"' and classid='1'  and pingtxt<>'' order by  id desc",conn,1,1
if not (rs.eof) then
Do While Not Rs.Eof
%>
<option value="<%=rs("pingtxt")%>"><%=rs("pingtxt")%></option>
<%
   Rs.MoveNext
   Loop
   End IF
%>
						    </select>    </TD>
  </TR>
  <TR>
    <TH><div align="right">备注：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
	<LABEL><INPUT name="bei" type="radio" id="bei"  value="马上收货好评" checked> 
	<SPAN class=font12l>马上收货好评</SPAN></LABEL>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;	<LABEL><INPUT type="radio" name="bei" id="bei"  value="一天后收货好评"> <SPAN class=font12l>一天后收货好评</SPAN></LABEL>(扣x*2个发布点)<BR>
    <LABEL><input type="radio" name="bei" id="bei" value="二天后收货好评"> <span class="font12l">二天后收货好评</span></LABEL>(扣x*2+1个发布点)
  <LABEL><input type="radio" name="bei" id="bei" value="三天后收货好评"> <span class="font12l">三天后收货好评</span></LABEL>(扣x*2+2个发布点)<BR>	</TD></TR>
  <TR>
    <TH><div align="right">接任务限制：</div></TH>
    <TD><select name="Limit" >
                                <option value="1" selected>不限制</option>
                                <option value="2">限制100积分以上</option>
                                <option value="3">限制300积分以上</option>
                                <option value="4">限制只能是VIP会员</option>
                                                                                          </select></TD>
  </TR>
  <TR>
    <TH><div align="right">卖家提醒语：</div></TH>
    <TD><INPUT   name=tips class=txtinput id=tips style="WIDTH: 290px" maxlength="100">
	<select name="lsx"  style="width:120px;" onChange="jieducm_tips(this.options[this.options.selectedIndex].value);"> 
								<option value="">请选择历史提醒语</option>
<%Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select top 10 tips from "&jieducm&"_zhongxin where username='"&session("useridname")&"' and classid='1'  and tips<>'' order by  id desc",conn,1,1
if not (rs.eof) then
Do While Not Rs.Eof
%>
									<option value="<%=rs("tips")%>"><%=rs("tips")%></option>
<%
   Rs.MoveNext
   Loop
   End IF
%>
						    </select></TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL><INPUT id=baohu disabled type=checkbox CHECKED value=1 name=baohu> 防来路保护 </LABEL>
      <LABEL><INPUT id=baohu3 disabled type=checkbox CHECKED value=1 name=baohu3> 防黄钻保护</LABEL>	 <LABEL> <input id=baohu32 disabled type=checkbox checked value=1  name=baohu32>自动检测宝贝地址和掌柜名</LABEL></TD></TR>
  
 <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL><INPUT name="Package"   type=checkbox id="Package"  disabled value=1 checked>
      套餐任务 </LABEL>
      (发布套餐任务请增加发布点)</TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><label><input  name="baohu2" type="checkbox" id="baohu2" value="1"  />  
                  任务保护，接手者接你任务后，要你审核才可以看到商品地址！</label></TD>
  </TR>
  <TR>
    <TH><div align="right">定时发布：</div></TH>
    <TD><input name="Timing" type="text"  onClick="SelectDate(this,'yyyy-MM-dd hh:mm:ss')"  readonly/>
                              留空则不生效</TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL><input  name="depot" type="checkbox" id="depot" value="1" />顺便添加到我的任务仓库</LABEL>&nbsp; 仓库备注(用于您自己分辨商品)： 
                               <label>
                               <input name="title" type="text" maxlength="20">
                               </label></TD></TR>
  
  <TR>
    <TD 
    style="PADDING-RIGHT: 0px; PADDING-LEFT: 200px; PADDING-BOTTOM: 5px; PADDING-TOP: 5px" 
    colSpan=2><input style="FONT-WEIGHT: bold; WIDTH: 120px; CURSOR: pointer; COLOR: #000000; HEIGHT: 26px" type="submit" name="button" id="button" <% if cunkuan<1 or fabudian<1 then%> disabled <%end if%> value="发布任务"  onClick="ajax()"><% if cunkuan<1 or fabudian<1 then%>存款或发布点低于1时不能再发任务<%end if%></TD></TR></FORM></TBODY></TABLE>
</DIV></DIV></DIV>    </td>
              </tr>
      </table>    </td>
              </tr>
      </table>	    </td>
	    </tr>
  </table>
  <script>
function ajax(){
document.form.button.disabled="disabled";
document.form.button.value="系统正在处理中";
document.form.submit();
}

  function jieducm_tips(s)
 {
var ss=s;
document.getElementById("tips").value=ss;

 }
   function jieducm_pingtxt(s)
 {
var ss=s;
document.getElementById("pingtxt").value=ss;
 }
 function jieducm_ping(num)
 {
 if (num==3)
	{document.getElementById("pinghistory").style.display="";}
	else
	{
		{document.getElementById("pinghistory").style.display="none";}

	}
 }
 
 function gouwuce(num)
 {	
 //alert(num);
	atuo(num);
 }
function atuo(num)
 {
	if (num==1)
	{
		document.getElementById("num").value=1;
		document.getElementById("jieducm_IfComeSet").disabled="";
		allnone();
		
	}
	else
	{
		//alert(num);
		document.getElementById("num").value=num;
		document.getElementById("jieducm_IfComeSet").disabled="disabled";
		document.getElementById("jieducm_IfComeSet").checked="";
		document.getElementById("showinfo_c1").style.display="none";
		document.getElementById("showinfo_c2").style.display="none";
		document.getElementById("showinfo_c3").style.display="none";
		allnone();
		for(var i=2;i<=num;i++)
		{
			document.getElementById("renwu"+i).style.display="";
		}
	}
 }
function allnone()
 {
	
	for( var i = 2;i <= 5;i++)
		{
			document.getElementById("renwu"+i).style.display="none";
		}
 }
 onload=function(){
// alert('ss');
 var nums=document.getElementById("num").value;
 //alert(nums);
 gouwuce(nums);

 }
 
 function setTaskid(num)
	  {
		  var str1="";
		  var str2="";
		  var tstr1="";
		  var tstr2="";
		  var ttstr="";
		  if(num==1)
		  {
			  str1="搜索店铺关键字";
			  str2="此处输入您需要告诉接手人商品在自己店铺内的什么地方可以找到，例如，进入店铺内搜索“100元面子卡”商品，第一个就是。";
			  tstr1="搜索提示";
			  tstr2="此处填写提示信息，说明店铺在关键字搜索结果列表中的位置，商品在店铺首页中大概位置等信息，比如：店铺在结果列表第二个，商品在商品首页第二排中间一个等等。";
			  //ttstr="是在淘宝搜索店铺进入您店里的流量";
		   }
		   else if(num==2)
		   {
			  str1="搜索宝贝关键字";
			  str2="推荐使用您的宝贝名称，如果您的宝贝名称在淘宝中重名商品过多，我们建议您修改宝贝名称后在设置此种来路任务或者使用店铺搜索的来路方式。";
			  tstr1="搜索提示";
			  tstr2="此处填写提示信息，说明任务商品在关键字搜索结果列表中的位置，比如第一页第7个。";
			  //ttstr="是在淘宝直接搜索宝贝进入的";
			}
		   else if(num==3)
		   {
			  str1="买家信用评价地址";
			  str2="此处输入已经购买过您的该任务宝贝买家的买家信用评价页面地址，可以在您的宝贝销售记录中点击买家用户名右侧的信用等级图标（红心，黄钻等）进入。";
			  tstr1="来路提示";
			  tstr2="此处填写提示信息，说明此宝贝购买记录在购买列表中的页数和位置等信息，比如：第二页倒数第一个。";
			  //ttstr="是从淘宝信用评价页面进入的，例如从买家评价的宝贝信息登陆您的宝贝";
		   }
		   window.tt1.innerHTML=str1;
		   window.ttt1.innerHTML=str2;
		   window.tt2.innerHTML=tstr1;
		   window.ttt2.innerHTML=tstr2;
		   //window.th1.innerHTML=ttstr;
		  
	  }
	  function SwitchDivDisplay(divName) {
		var obj_reg_info = document.getElementById(divName);
	   	if(obj_reg_info.style.display=="none") {
	       //	obj_reg_info.style.display="inline";
			document.getElementById("showinfo_c1").style.display="";
			document.getElementById("showinfo_c2").style.display="";
			document.getElementById("showinfo_c3").style.display="";
	   	} else {
	   		//obj_reg_info.style.display="block";
	   		//obj_reg_info.style.display="none";
			document.getElementById("showinfo_c1").style.display="none";
			document.getElementById("showinfo_c2").style.display="none";
			document.getElementById("showinfo_c3").style.display="none";
		}
	}
</script>
  <!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
  <%
  rs.close
  set rs=nothing
   call CloseConn()%>
</body>
</html>