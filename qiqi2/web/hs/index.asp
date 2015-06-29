<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="inc/index_conn.asp"-->
<!--#include file="jieducm_ad.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="config.asp"-->
<%
'请保留此声明信息,这段声明不并会影响你的速度!
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
<HTML lang=en><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta  name="description" content="<%=stupdesc%>"/>
<meta  name="keywords" content="<%=stupkey%>"/>
<LINK href="hs/main.css" type=text/css rel=stylesheet>
<LINK href="css/Css.css" type=text/css rel=stylesheet>
<link rel="stylesheet" id='skin' type="text/css" href="skin/qq/ymPrompt.css" />
<LINK href="css/top_bottom.css" type=text/css rel=stylesheet>
<LINK rel=stylesheet type=text/css href="css/global.css"> 
<SCRIPT language=JavaScript src="js/jquery.min.js"></SCRIPT>
<SCRIPT src="js/jieducm_pupu.js" type=text/javascript></SCRIPT>
<script type="text/javascript" src="js/ymPrompt.js"></script>
</HEAD>
<BODY  onload=document.form.name.focus()> 
<script language=javascript>
function CheckForm()
{
	if(document.form.name.value=="")
	{
		alert("请输入用户名！");
		document.form.name.focus();
		return false;
	}
	if(document.form.pass.value == "")
	{
		alert("请输入密码！");
		document.form.pass.focus();
		return false;
	}
	if (document.form.CheckCode)
	{
		if (document.form.CheckCode.value==""){
		   alert ("请输入您的验证码！");
		   document.form.CheckCode.focus();
		   return(false);
		}
	}
}
<%if stuppopup=1 then%>
	ymPrompt.alert({message:'<%=stupgonggao%>',title:'<%=stupkouhao%>!',height:250,width:400,fixPosition:true,dragOut:false});
		window.onload=function(){
			var o=document.getElementById('chgSkin');
			var css=document.getElementById('skin');
			o.selectedIndex=0;
			o.onchange=function(){
				css.href='skin/'+this.options[this.selectedIndex].value+'/ymPrompt.css';
			}
		}
		function json2str(o){
			var arr=[];
			var fmt=function(s){
				return /^(string|number)$/.test(typeof s)?"'"+s+"'":s;
			}
			for(var i in o) arr.push(i+':'+fmt(o[i]));
			return '{<br>&nbsp;&nbsp;'+arr.join(',<br>&nbsp;&nbsp;')+'&nbsp;&nbsp;<br>}';
		}
<%end if%>
</script>
<!--#include file=jieducm_top.asp-->
<%
sql="SELECT  sum(price) as jieducm_total FROM "&jieducm&"_zhongxin"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_total=rs("jieducm_total")
end if 

sql="SELECT  count(*) as jieducm_count FROM "&jieducm&"_zhongxin"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_count=rs("jieducm_count")
end if 
%>
<DIV class=page>
<DIV class=broadcast>
<DIV class=index-number style="OVERFLOW: hidden; WIDTH: 460px">
<span style="font-size:19px; color:#ff5e10" ><img src="images/jieducm_list.jpg" width="28" height="28" border="0">平台指数</span><SPAN>经手担保额：<EM><%=jieducm_total%></EM>元</SPAN><SPAN>任务总数：<EM><%=jieducm_count%></EM>个</SPAN>
</DIV>
<DIV class=real-time-info style="TEXT-ALIGN: center">
<MARQUEE scrollAmount=1 direction=up width=450 height=20>
<%
Sql = "select top 20 * from "&jieducm&"_zhongxin where  classid<>'6' order by id desc"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Rs.Eof Then
Response.Write("暂无信息")
Else
Do While Not Rs.Eof		
%>
<P style="MARGIN-BOTTOM: 10px"><EM><%=left(rs("username"),1)%>***<%=right(rs("username"),1)%> </EM>在 <SPAN><%=rs("now")%> 
</SPAN><EM>发布了 </EM>了一个担保金为<EM><%=rs("price")%></EM>元的任务。</P>
<%Rs.MoveNext
Loop
End IF%>
</MARQUEE></DIV></DIV></DIV>
<div class="main" style="height:226px;">
<div class="login" style="overflow:hidden;">
<div class="logint"><img src="hs/loginletter.gif" alt="平台最新公告" /></div>
<div class="loginm1">
         <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(2,10,15)%></TBODY></TABLE></TD></TR></TBODY></TABLE></td>
          </tr><div class="li2"></div>

</div>
</div>

<div class="mainads"><div class="centeradsm">
  <!--#include file="jieducm_flash.asp"-->
</div> 
</div>
</div>
<div class="tj">
<div class="middleBar">
  <div class="left"></div>
  <div class="right"></div>
  <div class="middle">
  
      <div class="howTo">
          <div class="title"><h2>淘宝刷信誉流程</h2><em>小提示：真实可靠的<b>淘宝刷信誉</b>、<b>淘宝刷钻</b>、<b>淘宝刷信用</b>虽然快，可千万不要贪心哦。</em></div>
          <div class="cont">
              <ol>
                <li class="step1">注册激活</li>
                <li class="step2">设操作码</li>
                <li class="step3">绑定买号<span>存款充值</span></li>
                <li class="step4">接手任务赚取发布点<span>使用存款购买发布点</span></li>
                <li class="step5">发布任务提升信誉度</li>
                <li class="step6">双方互评信誉</li>
              </ol>
          </div>
      </div>  
  </div>
</div></div>
<div class="main1">
<div class="inleft">
  <div class="helpnews">
<div class="helpnewst"><h3><table width="220" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="60" class="f12bt">新手入门</td>
                        <td width="170" align="right"><a href="newmore.asp?action=90">更多>></a></td>
                      </tr>
                  </table></h3></div>
<div class="helpnewsm"><ul>
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(90,8,15)%></TBODY></TABLE></TD></TR>
</ul>

</div>
</div></div>
<div class="inright">
<div class="helpteach" style="margin-top:0;">
<div class="helpteacht"><h3><table width="300" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="61" class="f12bt">卖家必读</td>
                        <td width="249" align="right"><a href="newmore.asp?action=92">更多>></a></td>
                      </tr>
                  </table></h3></div>
<div class="helpteachm">
<ul style="background:url(hs/jc1b.gif) no-repeat 220px center">
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(92,8,15)%></TBODY></TABLE></TD></TR>
</ul>
</div>
</div>
<div class="helpteach" style="margin-left:7px;margin-top:0;">
<div class="helpteacht"><h3><table width="300" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="61" class="f12bt">买家必读</td>
                        <td width="249" align="right"><a href="newmore.asp?action=93">更多>></a></td>
                      </tr>
                  </table></h3></div>
<div class="helpteachm">
<ul style="background:url(hs/jc2b.gif) no-repeat 220px center">
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(93,8,15)%></TBODY></TABLE></TD></TR>
</ul>
</div>
</div></div>
</div>
<div class="main1">
<div class="inleft">
  <div class="helpnews">
<div class="helpnewst"><h3><table width="220" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="60" class="f12bt">刷客必读</td>
                        <td width="170" align="right"><a href="newmore.asp?action=91">更多>></a></td>
                      </tr>
                  </table></h3></div>
<div class="helpnewsm"><ul>
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(91,8,15)%></TBODY></TABLE></TD></TR>
</ul>

</div>
</div></div>
<div class="inright">
<div class="helpteach" style="margin-top:0;">
<div class="helpteacht"><h3><table width="300" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="61" class="f12bt">赚钱窍门</td>
                        <td width="249" align="right"><a href="newmore.asp?action=95">更多>></a></td>
                      </tr>
                  </table></h3></div>
<div class="helpteachm">
<ul style="background:url(hs/jc1b.gif) no-repeat 220px center">
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(95,8,15)%></TBODY></TABLE></TD></TR>
</ul>
</div>
</div>
<div class="helpteach" style="margin-left:7px;margin-top:0;">
<div class="helpteacht"><h3><table width="300" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="61" class="f12bt">网店推广</td>
                        <td width="249" align="right"><a href="newmore.asp?action=93">更多>></a></td>
                      </tr>
                  </table></h3></div>
<div class="helpteachm">
<ul style="background:url(hs/jc2b.gif) no-repeat 220px center">
               <TR>
          <TD class=K_mtcontent 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          colSpan=6>
            <TABLE class=LeftNews cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
              <TBODY>
              <%call indexnews(94,8,15)%></TBODY></TABLE></TD></TR>
</ul>
</div>
</div></div>
</div>
<!--#include file=jieducm_bottom.asp-->
<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>