<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->
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
did=request.Form("did")
action=request.Form("fa")
if action="ok" then
Price=Replace(Trim(Request.Form("Price1")),"'","''")
ProUrl=Request.Form("ProUrl1")
Shopkeeper=Request.Form("Shopkeeper")
bei=Request.Form("bei")
baohu2=request.Form("baohu2")
isprice=request.form("isprice")
play=request.form("play")
Limit=request.Form("Limit")
title=Replace(Trim(Request.Form("title")),"'","''")
tips=Replace(Trim(request("tips")),"'","''")
addfabu=request.Form("addfabu")
if price="" then
	Response.Write("<script>alert('商品任务价不能为空!');history.back();</script>")
    response.End()
elseif int(price)<25  or int(price)>1000 then
	Response.Write("<script>alert('目前本平台只支持发布25-1000元的任务 ，请重新输入!');history.back();</script>")
    response.End()
elseif Shopkeeper="" then
	Response.Write("<script>alert('你还没有选择掌柜名!');history.back();</script>")
    response.End()	
elseif prourl="" then
	Response.Write("<script>alert('商品地址不能为空!');history.back();</script>")
    response.End()
end if

url= ProUrl
wstr=getHTTPPage(url) 
Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a class=""hCard fn"" href=""[^""]+""[^>]*>(.*?)</a>"
set Matches=objRegExp.Execute(wstr)
jieducm_sk=Matches(0).SubMatches(0)

Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a class=""hCard fn"" title=""(.*?)"""
set Matches=objRegExp.Execute(wstr)
jieducm_sk3=Matches(0).SubMatches(0)


if price>=25 and price<=40 then
   fabu=1
elseif price>40 and price<=100 then
   fabu=2
elseif price>100 and price<=200 then
   fabu=3
elseif price>200 and price<=500 then
  fabu=4
elseif price>500 and price<=1000 then
  fabu=5
end if 

if bei="一天后收货好评" then
fabu=fabu*2
elseif bei="二天后收货好评"  then
fabu=fabu*2+1
elseif bei="三天后收货好评"  then
fabu=fabu*2+2
end if
 

       Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_depot where id="&did&" and  username='"&session("useridname")&"'" ,Conn,3,3  
        if not rs.eof then
		rs("Price")=Price
    	rs("Shopkeeper")=Shopkeeper
		rs("ProUrl")=ProUrl
		rs("baohu")=baohu
		if baohu2="" then
		rs("baohu2")=0
		else
		rs("baohu2")=baohu2
		end if
		rs("now")=now
		rs("isprice")=isprice
		rs("play")=play
		rs("tips")=tips
		rs("Limit")=Limit
		rs("title")=title
		rs("addfabu")=addfabu
		rs("jieducm_fb")=fabu
    	rs.update
    	rs.close
		else
			Response.Write("<script>alert('操作有误!');history.back();</script>")
            response.End()
		end if
set rs=nothing
Response.Write("<script>alert('修改成功！');window.location.href='MySave.asp';</script>")
response.End()	
conn.close
set conn=nothing
end if


call myww(1)

id=request("id")
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_depot where id="&id&" and  username='"&session("useridname")&"'",Conn,1,1
if not rs.eof or not rs.bof then
Shopkeeper=rs("Shopkeeper")
bei=rs("bei")
baohu2=rs("baohu2")
isprice=Replace(Trim(rs("isprice")),"'","''")
play=rs("play")
tips=rs("tips")
Price=rs("price")
ProUrl=rs("ProUrl")
Limit=rs("Limit")
title=rs("title")
addfabu=rs("addfabu")
else
    Response.Write("<script>alert('出错了，请返回!');history.back();</script>")
	response.End()
end if
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
  <LI><A  href="pushmission.asp">发布任务</A> </LI>
  <LI><A href="ReMission.asp">已接任务</A> </LI>
  <LI><A href="MyMission.asp">已发任务</A> </LI>
  <LI><A  href="MyWw.asp">绑定店铺</A> </LI>
  <LI><A href="Buyhao.asp">绑定买号</A> </LI>
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="MySave.asp">任务仓库</A> </LI>
   <li><a href="../user/Express.asp">生成快递单</a> </li></UL></DIV>
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
修改任务仓库中任务须知： 
  <UL>
  <LI>1. 本功能是重新编辑任务详情；不会扣担保金和发布点。
  </LI>
  </UL>
</DIV></DIV><br>
<DIV class=addtask-wrap>
<DIV class=inner>
  <TABLE cellSpacing=0 cellPadding=0 border=0>
  <TBODY>
  <FORM name=form  action=""  method=post>
  <INPUT type=hidden value=ok name=fa> 
   <input type="hidden" name="did" value="<%=id%>">
  <TR>
    <TH><div align="right">商品任务价：</div></TH>
    <TD><input name=Price1 id=Price1 size="10" maxlength=6   onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=Price%>" >元注：(该价格是包含邮费的总价格)1-40元，扣一个发布点；41-100元扣两个发布点；101-200元，扣三个发布点；201-500元，扣四个发布点；501-1000元扣五个发布点</TD></TR>
  <TR>
    <TH><div align="right">掌&nbsp; 柜&nbsp; 名：</div></TH>
    <TD > <%	Sql = "select * from "&jieducm&"_keeper  where username='"&session("useridname")&"'  and class=1 "
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	IF Rs.Eof Then
	 Response.Write("暂无信息")
	Else
	Do While Not Rs.Eof
%><label><input  type='radio' <%if rs("keeper")=Shopkeeper then%> checked="checked" <%end if%> name='Shopkeeper'  id='Shopkeeper' value="<%=rs("keeper")%>"> <%=rs("keeper")%>
</label>
  <% Rs.MoveNext
   Loop
   End IF%></TD>
  </TR>
  <TR>
    <TH><div align="right">商品地址：</div></TH>
    <TD ><INPUT class=txtinput id=ProUrl1 style="WIDTH: 350px" maxLength=355 name=ProUrl1 value="<%=ProUrl%>"> 
    请填写商品页的正确网址</TD></TR>
  
 
    <tr>
      <TH><div align="right">增加发布点(套餐任务必选)：</div></TH>
      <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><select name="addfabu">
                <option value="0" <%if addfabu=0 then%> selected="selected" <%end if%>>不增加</option>
                <option value="1" <%if addfabu=1 then%> selected="selected" <%end if%>>增加1个</option>
                <option value="2" <%if addfabu=2 then%> selected="selected" <%end if%>>增加2个</option>
                <option value="3" <%if addfabu=3 then%> selected="selected" <%end if%>>增加3个</option>
                <option value="4" <%if addfabu=4 then%> selected="selected" <%end if%>>增加4个</option>
                <option value="5" <%if addfabu=5 then%> selected="selected" <%end if%>>增加5个</option>
              </select>         每多一件增加1个发布点，如发布点与套餐件数不符删除任务</TD>
    </TR>
    <TH><div align="right">价格是否相等：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><input name="isprice" type="radio" id="isprice" value="金额相等" <%if isprice="金额相等" then%> checked <%end if%>>
                              <span class="font12l"> 金额相等</span> <input type="radio" name="isprice" id="isprice" value="需修改价格" <%if isprice="需修改价格" then%> checked <%end if%>>
                               <span class="font12l">需修改价格</span>&nbsp; 任务价格和包括运费的商品总价格是否相等！</TD>
  </TR>
  <TR>
    <TH><div align="right">动态评分：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
	            <input name="play" type="radio" value="全部打5分"  <%if play="全部打5分" then%> checked <%end if%>>
                              <span class="font12l">全部打5分</span> <input type="radio" name="play" value="全部不打分"  <%if play="全部不打分" then%> checked <%end if%>>
                             <span class="font12l"> 全部不打分</span></TD>
  </TR>
   <TR>
    <TH></TH>
    <TD>&nbsp;</TD>
   </TR>
  <TR>
  
  <TR>
    <TH><div align="right">接任务限制：</div></TH>
    <TD><select name="Limit" >
                                <option value="1" <%if limit=1 then%> selected <%end if%>>不限制</option>
                                <option value="2" <%if limit=2 then%> selected <%end if%>>限制100积分以上</option>
                                <option value="3" <%if limit=3 then%> selected <%end if%>>限制300积分以上</option>
                                <option value="4" <%if limit=4 then%> selected <%end if%>>限制只能是VIP会员</option>
                                                                                          </select></TD>
  </TR>
  <TR>
    <TH><div align="right">卖家提醒语：</div></TH>
    <TD><INPUT   name=tips class=txtinput id=tips style="WIDTH: 290px" maxlength="100" value="<%=tips%>">
    发布后买家可见，100字以内</TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL><INPUT id=baohu disabled type=checkbox CHECKED value=1 name=baohu> 防来路保护 </LABEL>
      <LABEL><INPUT id=baohu3 disabled type=checkbox CHECKED value=1 name=baohu3> 防黄钻保护</LABEL>	  <input id=baohu32 disabled type=checkbox checked value=1 
      name=baohu32>
      自动检测宝贝地址和掌柜名</TD>
  </TR>
  
  <TR>
    <TH>&nbsp;</TH>
    <TD><input  name="baohu2" type="checkbox" id="baohu2" value="1" <%if baohu2=1 then%> checked="checked" <%end if%> />  
                  任务保护，接受者接你任务后，要你审核才可以看到商品地址！</TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL></LABEL>&nbsp; 仓库备注(用于您自己分辨商品)： 
                               <label>
                               <input name="title" type="text" maxlength="20" value="<%=title%>">
                               </label></TD></TR>
  
  <TR>
    <TD 
    style="PADDING-RIGHT: 0px; PADDING-LEFT: 200px; PADDING-BOTTOM: 5px; PADDING-TOP: 5px" 
    colSpan=2><input style="FONT-WEIGHT: bold; WIDTH: 120px; CURSOR: pointer; COLOR: #000000; HEIGHT: 26px" type="submit" name="button" id="button"  value="保存修改" onClick="this.disabled=true;document.form.submit();"></TD></TR></FORM></TBODY></TABLE>
</DIV></DIV></DIV>    </td>
              </tr>
      </table>	    </td>
	    </tr>
  </table>
  <!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
  <%
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing%>
</body>
</html>