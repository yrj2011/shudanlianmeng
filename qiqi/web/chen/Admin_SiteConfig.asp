<%@language=vbscript codepage=936 %>

<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
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
action=request("action")
if action="ok" then

Set rs1=server.createobject("ADODB.RECORDSET")
rs1.open "select * from "&jieducm&"_stup",conn,3,3
if not (rs1.eof) then
rs1("name")=request.Form("SiteName")
rs1("title")=request.Form("SiteTitle")
rs1("key")=request.Form("SiteKey")
rs1("desc")=request.Form("Sitedes")
rs1("gonggao")=request.Form("gonggao")
rs1("kouhao")=request.Form("kouhao")
rs1("url")=request.Form("SiteUrl")
rs1("logo")=request.Form("LogoUrl")
rs1("address")=request.Form("BannerUrl")
rs1("qq")=request.Form("qq")
rs1("qq2")=request.Form("qq2")
rs1("qq3")=request.Form("qq3")
rs1("qq4")=request.Form("qq4")
rs1("qq5")=request.Form("qq5")
rs1("qq6")=request.Form("qq6")
rs1("zfb")=request.Form("WebmasterName")
rs1("cft")=request.Form("cft")
rs1("email")=request.Form("WebmasterEmail")
rs1("phone")=request.Form("phone")
rs1("tel")=request.Form("tel")
rs1("icp")=request.Form("icp")
rs1("copy")=request.Form("Copyright")
rs1("fa")=request.Form("fa")
rs1("jifei")=request.Form("jifei")
rs1("kuan")=request.Form("kuan")
rs1("kuanvip")=request.Form("kuanvip")
rs1("kou")=request.Form("kou")
rs1("geshu")=request.Form("geshu")
rs1("zhang")=request.Form("zhang")
rs1("ping")=request.Form("ping")
rs1("zhu")=request.Form("zhu")
rs1("tjrjifei")=request.Form("tjrjifei")
rs1("tjrzhu")=request.Form("tjrzhu")
rs1("vipjifei")=request.Form("vipjifei")
rs1("fajifei")=request.Form("fajifei")
rs1("shouc")=request.Form("shouc")
rs1("shouf")=request.Form("shouf")
rs1("shouf2")=request.Form("shouf2")
rs1("shouxu")=request.Form("shouxu")
rs1("wu")=request.Form("wu")
rs1("shi")=request.Form("shi")
rs1("wushi")=request.Form("wushi")
rs1("yibai")=request.Form("yibai")
rs1("wubai")=request.Form("wubai")
rs1("ka")=request.Form("ka")
rs1("jifeidi")=request.Form("jifeidi")
rs1("jifeigao")=request.Form("jifeigao")
rs1("alllogin")=request.Form("alllogin")
rs1("fhc")=request.Form("fhc")
rs1("fhcshu")=request.Form("fhcshu")
rs1("quxin")=request.Form("quxin")
rs1("qushou")=request.Form("qushou")
rs1("quvip")=request.Form("quvip")
rs1("buyhao")=request.Form("buyhao")
rs1("xinshoufa")=request.Form("xinshoufa")
rs1("shouchangfa")=request.Form("shouchangfa")
rs1("vipfa")=request.Form("vipfa")
rs1("enchfa")=request.Form("enchfa")
rs1("jifeibig")=request.Form("jifeibig")
rs1("songfa")=request.Form("songfa")
rs1("songshou")=request.Form("songshou")
rs1("songji")=request.Form("songji")
rs1("car")=request.Form("car")
rs1("wangyin")=request.Form("wangyin")
rs1("alipay")=request.Form("alipay")
rs1("login")=request.Form("login2")
rs1("CheckCode")=request.Form("CheckCode")
rs1("isgive")=request.Form("isgive")
rs1("popup")=request.Form("popup")
rs1("register_taobao")=request.Form("register_taobao")
rs1("register_pai")=request.Form("register_pai")
rs1("register_you")=request.Form("register_you")
rs1("MailServer")=request.Form("MailServer")
rs1("MailServerUserName")=request.Form("MailServerUserName")
rs1("MailServerPassWord")=request.Form("MailServerPassWord")
rs1("MailDomain")=request.Form("MailDomain")
rs1("isemail")=request.Form("isemail")
rs1("ismessage")=request.Form("ismessage")
rs1("emailcontent")=request.Form("emailcontent")
rs1("messagecontent")=request.Form("messagecontent")
rs1("emailactive")=request.Form("emailactive")
rs1("puuser")=request.Form("puuser")
rs1("vipuser")=request.Form("vipuser")
rs1("tribestart")=request.Form("tribestart")
rs1("phonestart")=request.Form("phonestart")
rs1("jieweiok")=request.Form("jieweiok")
rs1("faweiok")=request.Form("faweiok")
rs1("invite")=request.Form("invite")
rs1("exitauto")=request.Form("exitauto")
rs1("vip_car_jieducm")=request.Form("vip_car_jieducm")
rs1("cang")=request.Form("cang")
rs1("kefu_pic")=request.Form("kefu_pic")
rs1("jieducm_gonggao")=request.Form("jieducm_gonggao")
rs1("jieducm_MerId")=request.Form("jieducm_MerId")
rs1("jieducm_MerKey")=request.Form("jieducm_MerKey")
rs1("yibao")=request.Form("yibao")
rs1("payis")=request.Form("payis")
rs1("payjifei")=request.Form("payjifei")
rs1("paynum")=request.Form("paynum")
rs1("vippaynum")=request.Form("vippaynum")
rs1("pupaynum")=request.Form("pupaynum")
rs1("RndQueryNum")=request.Form("RndQueryNum")
rs1("DefQueryNum")=request.Form("DefQueryNum")
rs1.update
rs1.close
set rs1=nothing
Response.Write("<script>alert('修改成功!');window.location.href='admin_siteconfig.asp';</script>")
end if
end if
	  
			  	Sql = "select* from "&jieducm&"_stup"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
  %>
<html>
<head>
<title>网站信息配置选顶</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" Class="border">
    <tr> 
     <td height="22" colspan=2 align=center class="title"><a name="Top"></a><b>网 站 配 置</b></td>
    </tr>
    <tr class="tdbg"> 
      <td width="70" height="30"><strong>管理导航：</strong></td>
      
    <td height="30"><a href="#SiteInfo">网站信息配置</a>| <a href="#Email">邮件服务器选项</a></td>
    </tr>
  </table>
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <form method="POST" action="Admin_SiteConfig.asp?action=ok" id="frm1" name="frm1"><tr><td height="25" class="tdbg"></td>
  <td width="942" height="25" class="tdbg">     </td>
    <tr>
      <td height="22" colspan="2" class="title"><strong>网站信息配置&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong><a href="#Top"><font color="#0066CC">（顶&nbsp;部）</font></a></td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>网站名称：</strong></td>
      <td width="942" height="25" class="tdbg"><input name="SiteName" type="text" id="SiteName" value="<%=rs("name")%>" size="40" maxlength="150">      </td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>网站标题：</strong></td>
      <td width="942" height="25" class="tdbg"><input name="SiteTitle" type="text" id="SiteTitle" value="<%=rs("title")%>" size="40" maxlength="250">      </td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>网站关键字：</strong><br>
        将被搜索引擎用来搜索您网站的关键内容<br>
        每个关键字用“|”号分隔</td>
      <td width="942" height="25" class="tdbg"><input name="SiteKey" type="text" id="SiteKey" value="<%=rs("key")%>" size="40" maxlength="250">      </td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>网站描述：</strong><br>
        将被搜索引擎用来说明您网站的主要内容<br>
        <font color="#FF6600">介绍中请不要带英文的逗号</font></td>
      <td width="942" height="25" class="tdbg"><input name="Sitedes" type="text" id="Sitedes" value="<%=rs("desc")%>" size="40" maxlength="250">      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>系统运行状态：</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="radio" name="invite" value="1"  <% if rs("invite")=1 then%> checked <% end if%>>
        正常运行      
        &nbsp;&nbsp;&nbsp;  <input type="radio" name="invite" value="2"  <% if rs("invite")=2 then%> checked <% end if%>>
        暂停维护
    </label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>是否关闭刷区：</strong></td>
      <td height="25" class="tdbg">
	    <input name="quxin" type="checkbox" id="quxin" value="1" <%if rs("quxin")=1 then%> checked <%end if%>>
        淘宝区 
          <input name="qushou" type="checkbox" id="qushou" value="1" <%if rs("qushou")=1 then%> checked <%end if%>>
          拍拍区          </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>是否关闭注册：</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="zhu" id="zhu" value="1"  <% if rs("zhu")="1" then%> checked <% end if%> >
        关闭注册 
          <input name="zhu" type="radio" id="zhu" value="2" <% if rs("zhu")="2" then%> checked <% end if%> >
          开启注册</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>注册会员是否需要邮件激活：</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="radio" name="emailactive" value="1"  <% if rs("emailactive")="1" then%> checked <% end if%> >需要
        <input type="radio" name="emailactive" value="2" <% if rs("emailactive")="2" then%> checked <% end if%>>不需要
      </label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>注册会员是否需要输入验证码：</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="CheckCode" id="CheckCode" value="1"  <% if rs("CheckCode")="1" then%> checked <% end if%> >
关闭验证码
  <input name="CheckCode" type="radio" id="CheckCode" value="2" <% if rs("CheckCode")="2" then%> checked <% end if%> >
开启验证码</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>注册会员是否需要审核：</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="login2" id="login2" value="1"  <% if rs("login")="1" then%> checked <% end if%> >
        关闭审核 
          <input name="login2" type="radio" id="login2" value="0" <% if rs("login")="0" then%> checked <% end if%> >
          开启审核</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>会员是否必须加入部落才能接发任务：</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="tribestart" id="tribestart" value="1"  <% if rs("tribestart")="1" then%> checked <% end if%> >
需要
  <input name="tribestart" type="radio" id="tribestart" value="0" <% if rs("tribestart")="0" then%> checked <% end if%> >
不需要</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>会员是否必须绑定手机才能接发任务：</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="phonestart" id="phonestart" value="1"  <% if rs("phonestart")="1" then%> checked <% end if%> >
需要
  <input name="phonestart" type="radio" id="phonestart" value="0" <% if rs("phonestart")="0" then%> checked <% end if%> >
不需要</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>新用户注册开通：</strong></td>
      <td height="25" class="tdbg">
	  <input name="register_taobao" type="checkbox" id="register_taobao" value="1" <%if rs("register_taobao")=1 then%> checked <%end if%>>
淘宝区
  <input name="register_pai" type="checkbox" id="register_pai" value="1" <%if rs("register_pai")=1 then%> checked <%end if%>>
拍拍区</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>注册会员是否可以赠送发布点：</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="isgive" id="isgive" value="1"  <% if rs("isgive")="1" then%> checked <% end if%> >
不可以
  <input name="isgive" type="radio" id="isgive" value="0" <% if rs("isgive")="0" then%> checked <% end if%> > 
  可以</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>是否关闭充值卡手动激活：</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="ka" id="ka" value="1"  <% if rs("ka")="1" then%> checked <% end if%> >
        关闭手动激活 
          <input name="ka" type="radio" id="ka" value="2" <% if rs("ka")="2" then%> checked <% end if%> >
          开启手动激活</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>是否关闭充值卡充值：</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="car" id="car" value="0" <% if rs("car")="0" then%> checked <% end if%>>
        关闭充值卡充值
        <input name="car" type="radio" id="car" value="1" <% if rs("car")="1" then%> checked <% end if%>>
        开启充值卡充值</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>是否关闭网银在线充值：</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="wangyin" id="wangyin" value="0" <% if rs("wangyin")="0" then%> checked <% end if%>>
        关闭网银在线充值
        <input type="radio" name="wangyin" id="wangyin" value="1" <% if rs("wangyin")="1" then%> checked <% end if%>>
        开启网银在线充值</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>是否关闭支付宝接口：</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="alipay" id="alipay" value="0" <% if rs("alipay")="0" then%> checked <% end if%>>
        关闭支付宝接口
          <input type="radio" name="alipay" id="alipay" value="1" <% if rs("alipay")="1" then%> checked <% end if%>>
        开启支付宝接口</td>
    </tr>
	
	<tr>
      <td height="25" class="tdbg"><strong>是否关闭易宝支付：</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="yibao" id="yibao" value="0" <% if rs("yibao")="0" then%> checked <% end if%>>
        关闭易宝支付
          <input type="radio" name="yibao" id="yibao" value="1" <% if rs("yibao")="1" then%> checked <% end if%>>
        开启易宝支付</td>
    </tr>
	
	
    <tr>
      <td height="25" class="tdbg"><strong>是否关闭所有用户登录：</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="alllogin" id="alllogin" value="1" <% if rs("alllogin")="1" then%> checked <% end if%>>
        限制所有用户登录 
          <input type="radio" name="alllogin" id="alllogin" value="0" <% if rs("alllogin")="0" then%> checked <% end if%>>
          恢复所有用户登录</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>不允许发布点兑换现金：</strong></td>
      <td height="25" class="tdbg"><input type="radio" name="fhc" id="fhc" value="1" <% if rs("fhc")="1" then%> checked <% end if%>>
        充许&nbsp; <input type="radio" name="fhc" id="fhc" value="0" <% if rs("fhc")="0" then%> checked <% end if%>>
        不充许</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>新用户注册送：</strong></td>
      <td height="25" class="tdbg">发布点
        <input name="songfa" type="text" id="songfa" value="<%=rs("songfa")%>" onKeyUp="if(isNaN(value))execCommand('undo')">
        个，积分：
        <input name="songji" type="text" id="songji" value="<%=rs("songji")%>" onKeyUp="if(isNaN(value))execCommand('undo')">
        个</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>超时自动退出任务多少次不能再接任务：</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="text" name="exitauto" onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=rs("exitauto")%>">
      次</label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>同一个会员只能接手多少个未完成的任务：</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="text" name="jieweiok" onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=rs("jieweiok")%>">
        个
      </label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>同一个会员只能发布多少个未完成的任务：</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="text" name="faweiok" onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=rs("faweiok")%>">
      个</label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>一个账号同时允许绑定多少个淘宝买号：</strong></td>
      <td height="25" class="tdbg"><input type="text" name="buyhao" id="buyhao" size="40" maxlength="250" value="<%=rs("buyhao")%>" onKeyUp="if(isNaN(value))execCommand('undo')"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>多少个发布点充许兑换现金：</strong></td>
      <td height="25" class="tdbg"><input name="fhcshu" type="text" id="fhcshu" value="<%=rs("fhcshu")%>" size="40" maxlength="250" onKeyUp="if(isNaN(value))execCommand('undo')"></td>
    </tr>
    
    
    <tr>
      <td height="25" class="tdbg"><strong>是否弹出平台公告：</strong></td>
      <td height="25" class="tdbg">
	  <select name="popup">
        <option value="1" <%if rs("popup")=1 then%> selected="selected" <%end if%>>弹出</option>
        <option value="2" <%if rs("popup")=2 then%> selected="selected" <%end if%>>关闭</option>
      </select>      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>网站弹出公告：</strong></td>
      <td height="25" class="tdbg"><input name="gonggao" type="text" id="gonggao" value="<%=rs("gonggao")%>" size="40" maxlength="250"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>网站弹出公告标题：</strong></td>
      <td height="25" class="tdbg"><input name="kouhao" type="text" id="kouhao" value="<%=rs("kouhao")%>" size="40" maxlength="250"></td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>网站地址：</strong><br>
        请添写完整URL地址</td>
      <td width="942" height="25" class="tdbg"><input name="SiteUrl" type="text" id="SiteUrl" value="<%=rs("url")%>" size="40" maxlength="255"> 
      (后面不要加/)      </td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>LOGO地址：</strong><br>
        请添写完整URL地址</td>
      <td width="942" height="25" class="tdbg"><input name="LogoUrl" type="text" id="LogoUrl" value="<%=rs("logo")%>" size="40" maxlength="255">
          <input type="button" name="Submit" value="上传图片" onClick="window.open('../upload_flash.asp?formname=frm1&editname=LogoUrl&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
        (logo大小175*68px) </td>
    </tr>
	 <tr>
      <td height="25" class="tdbg"><strong>首页客服图片：</strong></td>
      <td height="25" class="tdbg"><input name="kefu_pic" type="text" id="kefu_pic" value="<%=rs("kefu_pic")%>" size="40" maxlength="255">
      <input type="button" name="Submit2" value="上传图片" onClick="window.open('../upload_flash.asp?formname=frm1&editname=kefu_pic&uppath=uploadpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
      (大小253*130px) </td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>公司地址：</strong><br></td>
      <td width="942" height="25" class="tdbg"><input name="BannerUrl" type="text" id="BannerUrl" value="<%=rs("address")%>" size="40" maxlength="255">      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>客服热线：</strong></td>
      <td height="25" class="tdbg"><input name="phone" type="text" id="phone" value="<%=rs("phone")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>咨询热线：</strong></td>
      <td height="25" class="tdbg"><input name="tel" type="text" id="tel" value="<%=rs("tel")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>在线咨询QQ：</strong></td>
      <td height="25" class="tdbg"><input name="qq5" type="text" id="qq5" value="<%=rs("qq5")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>新手客服QQ：</strong></td>
      <td height="25" class="tdbg"><input name="qq" type="text" id="qq" value="<%=rs("qq")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>充值客服QQ：</strong></td>
      <td height="25" class="tdbg"><input name="qq2" type="text" id="qq2" value="<%=rs("qq2")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>提现客服QQ：</strong></td>
      <td height="25" class="tdbg"><input name="qq3" type="text" id="qq3" value="<%=rs("qq3")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>投诉客服QQ：</strong></td>
      <td height="25" class="tdbg"><input name="qq4" type="text" id="qq4" value="<%=rs("qq4")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>广告合作QQ：</strong></td>
      <td width="942" height="25" class="tdbg"><input name="qq6" type="text" id="qq6" value="<%=rs("qq6")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>财富通账号：</strong></td>
      <td height="25" class="tdbg"><input name="cft" type="text" id="cft" value="<%=rs("cft")%>" size="40" maxlength="200"></td>
    </tr>
    
    <tr>
      <td height="25" class="tdbg"><strong>被推荐人购买vip卡奖励多少个发布点：</strong></td>
      <td height="25" class="tdbg"><input name="vip_car_jieducm" type="text" id="vip_car_jieducm" value="<%=rs("vip_car_jieducm")%>" size="40" maxlength="200"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>接（发）一个任务可得多少积分：</strong></td>
      <td height="25" class="tdbg"><input name="fajifei" type="text" id="fajifei" value="<%=rs("fajifei")%>" size="40" maxlength="200" onKeyUp="if(isNaN(value))execCommand('undo')"></td>
    </tr>
   
    <tr>
      <td height="25" class="tdbg"><strong>被推荐人注册时推荐人得多少积分：</strong></td>
      <td height="25" class="tdbg"><input name="tjrzhu" type="text" id="tjrzhu" value="<%=rs("tjrzhu")%>" size="40" maxlength="200" onKeyUp="if(isNaN(value))execCommand('undo')"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>被推荐人接（发）一个任务推荐人得多少积分：</strong></td>
      <td height="25" class="tdbg"><input name="tjrjifei" type="text" id="tjrjifei" value="<%=rs("tjrjifei")%>" size="40" maxlength="200" onKeyUp="if(isNaN(value))execCommand('undo')"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>一个月接手同一个人任务数：</strong></td>
      <td height="25" class="tdbg"><input name="geshu" type="text" id="geshu" value="<%=rs("geshu")%>" size="40" maxlength="20" onKeyUp="if(isNaN(value))execCommand('undo')">
        个</td>
    </tr>
	<tr>
      <td height="25" class="tdbg"><strong>发布一个收藏任务扣多少发布点：</strong></td>
      <td height="25" class="tdbg"><input name="cang" type="text" id="cang" value="<%=formatnumber(rs("cang"),2,-1)%>" size="40" maxlength="20" onKeyUp="if(isNaN(value))execCommand('undo')"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>积分换发布点比例：</strong></td>
      <td height="25" class="tdbg"><input name="jifei" type="text" id="jifei" value="<%=rs("jifei")%>" size="40" maxlength="20" onKeyUp="if(isNaN(value))execCommand('undo')">
        (必须是数字,多少个积分可以换一个发布点)</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>发布点回收价格：</strong></td>
      <td height="25" class="tdbg"><input name="fa" type="text" id="fa" value="<%=formatnumber(rs("fa"),2,-1)%>" size="40" maxlength="20" onKeyUp="if(isNaN(value))execCommand('undo')">
        (必须是数字,1个发布点多少钱)</td>
    </tr>
	
    <tr>
      <td height="25" class="tdbg"><strong>发布点出售价格：</strong></td>
      <td height="25" class="tdbg">普通会员价格：
        <input name="kuan" type="text" id="kuan" value="<%=formatnumber(rs("kuan"),2,-1)%>" size="10" maxlength="20" onKeyUp="if(isNaN(value))execCommand('undo')">
        元
        &nbsp;&nbsp; VIP会员价格：
        <input name="kuanvip" type="text" id="kuanvip" value="<%=formatnumber(rs("kuanvip"),2,-1)%>" size="10" maxlength="20" onKeyUp="if(isNaN(value))execCommand('undo')">
       元       (必须是数字,每个发布点多少钱)</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>接任务时积分高于：</strong>(接任务得发布点比例)</td>
      <td height="25" class="tdbg"><input type="text" name="jifeibig" id="jifeibig" value="<%=rs("jifeibig")%>">
        个积分时，得发布点
          <input name="kou" type="text" id="kou"  value="<%=formatnumber(rs("kou"),1,-1)%>"  size="20" maxlength="20" onKeyUp="if(isNaN(value))execCommand('undo')">
         
(必须是数字可以为小数如:0.9)如果低于设置积分得发布点比例：1：1</td>
    </tr>
    
     <tr>
      <td height="25" class="tdbg"><strong>发送手机短信收费标准：</strong></td>
      <td height="25" class="tdbg"><label>
        普通用户一条<input name="puuser" value="<%=rs("puuser")%>" type="text" size="10" maxlength="10" onKeyUp="if(isNaN(value))execCommand('undo')">分
        vip用户一条<input name="vipuser" value="<%=rs("vipuser")%>" type="text" size="10" maxlength="10" onKeyUp="if(isNaN(value))execCommand('undo')">分
      </label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>VIP会员接任务时得多少发布点比例：</strong></td>
      <td height="25" class="tdbg"><input type="text" name="vipjifei" id="vipjifei" value="<%=formatnumber(rs("vipjifei"),1,-1)%>">        (必须是数字可以为小数如:0.9)</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>赠送发布点手继费：</strong></td>
      <td height="25" class="tdbg"><input name="shouxu" type="text" id="shouxu" value="<%=rs("shouxu")%>" size="40" maxlength="100" onKeyUp="if(isNaN(value))execCommand('undo')">
      &nbsp; (必须是数字,例如输入10,就是收取10%的手继费)</td>
    </tr>
    <tr>
      <td width="400" height="25" class="tdbg"><strong>站长信箱：</strong></td>
      <td width="942" height="25" class="tdbg"><input name="WebmasterEmail" type="text" id="WebmasterEmail" value="<%=rs("email")%>" size="40" maxlength="100">      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>网银商户号：</strong></td>
      <td height="25" class="tdbg"><input name="wu" type="text" id="wu" value="<%=rs("wu")%>" size="40" maxlength="100"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>网银MD5密钥：</strong></td>
      <td height="25" class="tdbg"><input name="shi" type="text" id="shi" value="<%=rs("shi")%>" size="40" maxlength="100"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>支付宝账号：</strong></td>
      <td height="25" class="tdbg"><input name="WebmasterName" type="text" id="WebmasterName" value="<%=rs("zfb")%>" size="40" maxlength="200"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>支付宝账号对应的partnerID：</strong></td>
      <td height="25" class="tdbg"><input name="wushi" type="text" id="wushi" value="<%=rs("wushi")%>" size="40" maxlength="100"></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>支付宝安全校验码：</strong></td>
      <td height="25" class="tdbg"><input name="yibai" type="text" id="yibai" value="<%=rs("yibai")%>" size="40" maxlength="100"></td>
    </tr>
    
    <tr>
      <td height="25" class="tdbg"><strong>易宝商户编号：</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="text" name="jieducm_MerId" size="40" maxlength="100" value="<%=rs("jieducm_MerId")%>">
      </label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>易宝密钥：</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="text" name="jieducm_MerKey" size="40" maxlength="100" value="<%=rs("jieducm_MerKey")%>">
      </label></td>
    </tr>
	<tr>
      <td height="25" class="tdbg"><strong>是否开通发布点出售功能：</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="radio" name="payis" value="1" <%if rs("payis")=1 then%> checked <%end if%>>开通
        <input type="radio" name="payis" value="0" <%if rs("payis")=0 then%> checked <%end if%>>关闭
      </label></td>
    </tr>
	    <tr>
      <td height="25" class="tdbg"><strong>出售发布点最低积分为：</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="text" name="payjifei" onKeyUp="if(isNaN(value))execCommand('undo')"  value="<%=rs("payjifei")%>">
      </label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>出售发布点最低数量为：</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="text" name="paynum" onKeyUp="if(isNaN(value))execCommand('undo')"  value="<%=rs("paynum")%>">
      个</label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>出售发布点手续费为：</strong></td>
      <td height="25" class="tdbg">VIP会员：
        <label>
        <input name="vippaynum" type="text" size="10" onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=rs("vippaynum")%>">
     %&nbsp;&nbsp; 普通会员：
     <input name="pupaynum" type="text" size="10" onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=rs("pupaynum")%>">
      %</label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>查询快递单号使用次数：</strong></td>
      <td height="25" class="tdbg"><label>
        随机生成
        <input type="text" name="RndQueryNum" size="10" onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=rs("RndQueryNum")%>">
        次   
       &nbsp; 自定义生成
        <input type="text" name="DefQueryNum" size="10" onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=rs("DefQueryNum")%>">
      次</label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>顶部公告：</strong></td>
      <td height="25" class="tdbg"><label>
        <textarea name="jieducm_gonggao" cols="40" rows="4" id="jieducm_gonggao"><%=rs("jieducm_gonggao")%></textarea>
      </label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>流量统计代码：<br>
      </strong>也可加入QQ咨询台代码</td>
      <td height="25" class="tdbg">
	  <textarea name="icp" cols="40" rows="4" id="icp"><%=rs("icp")%></textarea>
      &nbsp;</td>
    </tr>

     <tr> 
      <td height="25" colspan="2" class="title"><a name="Email"></a><strong>邮件服务器选项&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong><a href="#Top"><font color="#0066CC">（顶&nbsp;部）</font></a></td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>邮件发送组件：</strong><br>
        请一定要选择服务器上已安装的组件</td>
      <td height="25" class="tdbg"> 
        (服务器须 JMAIL 支持)</td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>SMTP服务器地址：</strong><br>
        用来发送邮件的SMTP服务器<br>
        如果你不清楚此参数含义，请联系你的空间商 </td>
      <td height="25" class="tdbg"> 
        <input name="MailServer" type="text" id="MailServer" value="<%=rs("MailServer")%>" size="40">      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>SMTP登录用户名：</strong><br>
        当你的服务器需要SMTP身份验证时还需设置此参数</td>
      <td height="25" class="tdbg"> 
        <input name="MailServerUserName" type="text" id="MailServerUserName" value="<%=rs("MailServerUserName")%>" size="40">      </td>
    </tr>
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>SMTP登录密码：</strong><br>
        当你的服务器需要SMTP身份验证时还需设置此参数 </td>
      <td height="25" class="tdbg"> 
        <input name="MailServerPassWord" type="PassWord" id="MailServerPassWord" value="<%=rs("MailServerPassWord")%>" size="40">      </td>
    </tr>
    
    <tr> 
      <td width="400" height="25" class="tdbg"><strong>SMTP域名</strong>：<br>
        如果用“xieshanhui@163.com”这样的用户名登录时，请指明778892.com</td>
      <td height="25" class="tdbg"> 
        <input name="MailDomain" type="text" id="MailDomain" value="<%=rs("MailDomain")%>" size="40">      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>新用户注册是否发E_mail：</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="checkbox" name="isemail" value="1" <%if rs("isemail")=1 then%> checked="checked" <%end if%>>
      选择则用户注册后将发送欢迎邮件</label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>新用户注册是否发站内消息：</strong></td>
      <td height="25" class="tdbg"><label>
        <input type="checkbox" name="ismessage" value="1" <%if rs("ismessage")=1 then%> checked="checked" <%end if%>>
      选择则新用户注册发送站内消息</label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>Email内容设置：</strong></td>
      <td height="25" class="tdbg"><label>
        <textarea name="emailcontent" cols='90' rows='6'><%=rs("emailcontent")%></textarea>
      </label></td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><strong>站内消息内容：</strong></td>
      <td height="25" class="tdbg"><textarea name="messagecontent" cols='90' rows='6'><%=rs("messagecontent")%></textarea></td>
    </tr>
    <tr>
      <td height="40" colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="SaveConfig">
          <input name="cmdSave" type="submit" id="cmdSave" value=" &nbsp;保存设置&nbsp; " <% 'If ObjInstalled=false Then response.write "disabled" %> style="cursor: hand;background-color: #cccccc;">      </td>
    </tr>
</td>
      </tr>
  </form>
  </table>
<%end if%>
  <!--#include file="Admin_fooder.asp" -->
</body>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</html>
