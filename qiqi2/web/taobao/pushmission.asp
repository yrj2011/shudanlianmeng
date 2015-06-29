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
if session("czm")="" then
 response.Redirect("code.asp")
 response.end()
elseif vip="1" then
 response.Redirect("pushmissionvip.asp")
 response.End()
end if
call myww(1)

%>
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh" lang="zh" dir="ltr">
<head>
<title>会员中心-<%=stupname%> -<%=stuptitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="../css/style.css" type=text/css rel=stylesheet>
<LINK href="../css/base.css" type=text/css rel=stylesheet>
<LINK href="../css/layout.css" type=text/css rel=stylesheet>
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
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="pushmission.asp">发布任务</A> </LI>
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
              <DIV class=page>
<DIV class=addtask-wrap>
<DIV class=inner>
发布任务前须知： 
<UL>
  <LI>1. 发布一个任务会扣除您的存款，请保证您的平台存款充足； 
  <LI>2. 请保证您在本站有足够的存款，您所发布任务价格金额从存款里面扣除，有人接到任务后做作为商品款返回给你。 </LI>
  <li>3. 请不要帮别人发布任务，否则会导致您的朋友无法在本站发布。</li>
  </UL></DIV></DIV><br>
<DIV class=addtask-wrap>
<DIV class=inner>
  <TABLE cellSpacing=0 cellPadding=0 border=0>
  <TBODY>
  <FORM name=form id="form"  action="jieducm_faok.asp"  method=post>
  <INPUT type=hidden value=ok name=fa> 
  <TR>
    <TH><div align="right">商品任务价：</div></TH>
    <TD><input name=Price1 id=Price1 size="10" maxlength=6   onKeyUp="if(isNaN(value))execCommand('undo')" >元注：(该价格是包含邮费的总价格)1-40元，扣一个发布点；41-100元扣两个发布点；101-200元，扣三个发布点；201-500元，扣四个发布点；501-1000元扣五个发布点</TD></TR>
  <TR>
    <TH><div align="right">掌&nbsp; 柜&nbsp; 名：</div></TH>
    <TD ><%call shopname(1)%></TD>
  </TR>
  <TR>
    <TH><div align="right">商品地址：</div></TH>
    <TD ><INPUT class=txtinput id=ProUrl1 style="WIDTH: 350px" maxLength=355 name=ProUrl1>
    请填写商品页的正确网址</TD></TR>
	   <TR>
    <TH><div align="right">是否来路搜索任务：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><label>
      <input type="checkbox" name="jieducm_IfComeSet" value="1" onClick="javascript:SwitchDivDisplay('showinfo_c1')" >
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
 
    <TH><div align="right">价格是否相等：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">  <label><input name="isprice" type="radio" id="isprice" value="金额相等" checked>
    <span class="font12l">金额相等</span> </label>  
    <label><input type="radio" name="isprice" id="isprice" value="需修改价格">
                              <span class="font12l">需修改价格</span></label>
                              &nbsp; 任务价格和包括运费的商品总价格是否相等！</TD>
  </TR>
  <TR>
    <TH><div align="right">动态评分：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><label><input name="play" type="radio" value="全部打5分" checked>
                              <span class="font12l">全部打5分</span></label><label> <input type="radio" name="play" value="全部不打分">
                             <span class="font12l"> 全部不打分</span></label></TD>
  </TR>
   <TR>
    <TH></TH>
    <TD><FONT color=#ff0000><STRONG>* 
      发布延时收货的任务，平台免费提供物流单号，并强制买家延时收货</STRONG></FONT></TD></TR>
  <TR>
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
						    </select>
    </TD>
  </TR>
  <TR>
    <TH><div align="right">备注：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
	<LABEL><INPUT name="bei" type="radio" id="bei"  value="马上收货好评" checked> 
	<SPAN class=font12l>马上收货好评</SPAN></LABEL>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;	<LABEL><INPUT type="radio" name="bei" id="bei"  value="一天后收货好评"> <SPAN class=font12l>一天后收货好评</SPAN></LABEL>(扣x*2个发布点)<BR>
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
      <LABEL><INPUT id=baohu3 disabled type=checkbox CHECKED value=1 name=baohu3> 防黄钻保护</LABEL>	<LABEL>  <input id=baohu32 disabled type=checkbox checked value=1 name=baohu32> 自动检测宝贝地址和掌柜名</LABEL></TD>
  </TR>
  
  <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL><INPUT id="Package" disabled type=checkbox  value=1 name="Package">
      套餐任务 </LABEL>
      (VIP会员可以发布套餐任务的宝贝)</TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><label><input  name="baohu2" type="checkbox" id="baohu2" value="1"  />  
                  任务保护，接手者接你任务后，要你审核才可以看到商品地址！</label></TD>
  </TR>
    <TR>
    <TH><div align="right">定时发布：</div></TH>
    <TD> <input name="Timing" type="text"  onClick="SelectDate(this,'yyyy-MM-dd hh:mm:ss')"  readonly/>
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
      </table>	    </td>
	    </tr>
  </table>
<script>
function ajax(){
document.form.button.disabled="disabled";
document.form.button.value="系统正在处理中";
document.form.submit();
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

</script>
  <!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
  <%
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing%>
</body>
</html>