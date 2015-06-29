<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
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
'Copyright (C) 2008－2009 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
did=request.Form("did")
action=request.Form("fa")
if action="ok" then
ProUrl=Request.Form("ProUrl")
Shopkeeper=Replace(Trim(Request.Form("Shopkeeper")),"'","''")
bei=Replace(Trim(Request.Form("info[remarks]")),"'","''")
product=request.form("product[number]")
title=Replace(Trim(Request.Form("title")),"'","''")
tips=Replace(Trim(request("tips")),"'","''")

if Shopkeeper="" then
	Response.Write("<script>alert('你还没有选择掌柜名!');history.back();</script>")
    response.End()	
elseif prourl="" then
	Response.Write("<script>alert('商品地址不能为空!');history.back();</script>")
    response.End()
elseif bei="" then
	Response.Write("<script>alert('你还没有选择备注!');history.back();</script>")
    response.End()
elseif product="" then
	Response.Write("<script>alert('收藏数量不能为空!');history.back();</script>")
    response.End()
elseif product<1 then
	Response.Write("<script>alert('收藏数量不能小于1!');history.back();</script>")
    response.End()
elseif not isNumeric(product) then
  Response.Write("<script>alert('输入有误请重新输入！');history.back();</script>")
  response.End()
end if

url= ProUrl
wstr=getHTTPPage(url) 
Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a class=""hCard fn"" href=""[^""]+""[^>]*>(.*?)</a>"
set Matches=objRegExp.Execute(wstr)
boyu=Matches(0).SubMatches(0)

Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a href="".*?"" class=""hCard fn"">(.*?)</a>"
set Matches=objRegExp.Execute(wstr)
boyu2=Matches(0).SubMatches(0)

if (boyu=shopkeeper) or (boyu2=shopkeeper) then
else
	Response.Write("<script>alert('此掌柜名与填写的商品地址不符!请重新输入!');history.back();</script>")
    response.End()
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_depot where id="&did&" and  username='"&session("useridname")&"' " ,Conn,3,3  
if not rs.eof then
rs("Shopkeeper")=Shopkeeper
rs("ProUrl")=ProUrl
rs("username")=session("useridname")
rs("now")=now
rs("bei")=bei
rs("pid")=pid
rs("num2")=product
rs("title")=title
rs("tips")=tips
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
num2=rs("num2")
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
<SCRIPT language=JavaScript src="../js/jquery.min.js"></SCRIPT>

</head>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=../taobao/jieducm_toptao.asp-->
<SCRIPT type=text/javascript>
  function changPrice()
  {
  var price, number, flag;
  number = document.getElementById('product[number]').value;
  if   ((number.indexOf("-")   ==   0)||!(number.indexOf(".")   ==   -1)){   
  document.getElementById('product[number]').value=""
  alert("收藏数量不能为小数或负数");   
  return   false;   
  }
number = number*(<%=stupcang%>*10)/10;
  $('#collect_price').html(number);
  }
  
  </SCRIPT>
<div align="center">  
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 900px">
<UL id=task_nav>
 <li><a  href="index.asp">淘宝收藏区</a> </li>
  <LI><A  href="pushmission.asp">发收藏任务</A> </LI>
  <LI><A href="ReMission.asp">已接任务</A> </LI>
  <LI><A href="MyMission.asp">已发任务</A> </LI>
  <LI><A  href="MyWw.asp">绑定店铺</A> </LI>
  <LI><A href="Buyhao.asp">绑定买号</A> </LI>
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="MySave.asp">任务仓库</A> </LI>
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
修改任务仓库中任务须知： 
  <UL>
  <LI>1. 本功能是重新编辑任务详情；不扣发布点。  </LI>
  </UL>
</DIV></DIV><br>
<DIV class=addtask-wrap>
<DIV class=inner>
  <TABLE cellSpacing=0 cellPadding=0 border=0>
  <TBODY>
  <FORM name=form  action=""  method=post id="form">
  <INPUT type=hidden value=ok name=fa> 
   <input type="hidden" name="did" value="<%=id%>">
  
  <TR>
    <TH><div align="right">掌&nbsp; 柜&nbsp; 名：</div></TH>
    <TD > <%	Sql = "select * from "&jieducm&"_keeper  where username='"&session("useridname")&"'  and class=1 "
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	IF Rs.Eof Then
	 Response.Write("暂无信息")
	Else
	Do While Not Rs.Eof
%><label><input  type='radio' <%if rs("keeper")=Shopkeeper then%> checked="checked" <%end if%> name='Shopkeeper'  id='nick_name' value="<%=rs("keeper")%>"> <%=rs("keeper")%>
</label>
  <% Rs.MoveNext
   Loop
   End IF%></TD>
  </TR>
  
    <TH><div align="right">备注：</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
	<input name="info[remarks]" type="radio" id="info[remarks]" value="宝贝收藏任务" <%if bei="宝贝收藏任务" then%> checked <%end if%>>
                              <span class="font12l"> 宝贝收藏任务</span> <input type="radio" name="info[remarks]" id="info[remarks]" value="店铺收藏任务" <%if bei="店铺收藏任务" then%> checked <%end if%>>
                               <span class="font12l">店铺收藏任务</span>&nbsp;</TD>
  </TR>
    
<TR>
    <TH><div align="right">请输入店铺或宝贝地址：</div></TH>
    <TD ><INPUT class=txtinput id=ProUrl style="WIDTH: 165px" maxLength=355 name=ProUrl value="<%=prourl%>"> 
    宝贝收藏任务只能发布宝贝链接地址，店铺收藏任务只能发布店铺地址</TD>
  </TR>
    <TR>
    <TH><div align="right">收藏数量：</div></TH>
    <TD style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"> 
      <label>
        <INPUT id=product[number] maxLength=255 onchange=changPrice() name=product[number] onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=num2%>"> 
        </label>
     
   <SPAN class="blue bold" style="FONT-SIZE: 14px">该任务共需<SPAN style="color:#FF0000" id=collect_price>0</SPAN>个发布点</SPAN> </TD>
  </TR><TR>
    <TH><div align="right">收藏标签：</div></TH>
    <TD><INPUT id=tips maxLength=10 name=tips value="<%=tips%>">
      不要超过8个汉字</TD>
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