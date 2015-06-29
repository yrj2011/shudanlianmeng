<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<%
'------------------------------------------------------------------
'***************************拍拍信誉互刷系统***********************
'本程序专为拍拍免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：捷度传媒
'官方主页：http://www.jieducm.cn
'技术支持：robin 
'QQ;616591415
'电话：13598006137
'版权：版权属于捷度传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2008－2011 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
action=request("action")
if action="ok" then
if czm<>request("czm") then
   Response.Write("<script>alert('操作码有误请重新输入！');history.back();</script>")
   response.End()
end if
hao=Replace(Trim(Request.Form("hao")),"'","''")
	Sql = "select count(*) as xinyu from "&jieducm&"_xinyu where username='"&session("useridname")&"' and class=2 and start=0 "
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	IF not Rs.Eof Then
	 ii=rs("xinyu")
	End IF
				
       if ii>=stupbuyhao then
	   	 Response.Write("<script>alert('一个号只能绑定"&stupbuyhao&"个买号!');history.back();</script>")
         response.End()
	   end if
	   
      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_xinyu where shopname='"&request("hao")&"'" ,Conn,3,3  
	  if rs.eof then
	    rs.addnew
		rs("username")=session("useridname")
		rs("shopname")=hao
		rs("prourl")=pingjia
		rs("num")=10
		rs("class")=2
		rs("start")=0
    	rs.update
    	rs.close
		 Response.Write("<script>alert('绑定成功!');window.location.href='buyhao.asp';</script>")
		 response.End()
	 else
	 	 Response.Write("<script>alert('此号已被其它用户绑定!');history.back();</script>")
		 response.End()
	 end if
elseif action="del" then
id=request.QueryString("id")
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_xinyu where id="&id&" and username='"&session("useridname")&"'",conn,3,3
rs.close
Response.Write("<script>window.location.href='buyhao.asp';</script>")
response.End()
end if
%>	
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<SCRIPT src="../js/Dialog.js" type=text/javascript></SCRIPT>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=../taobao/jieducm_toptao.asp-->

<!-- 左隐藏菜单结束 --><!-- 左隐藏菜单BIG -->
<div align="center">
<DIV 
style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 700px">
<UL id=paipai_task_nav>
  <LI><A  href="index.asp">拍拍互动区</A> </LI>
  <LI><A   href="pushmission.asp">发布任务</A> </LI>
  <LI><A href="ReMission.asp">已接任务</A> </LI>
  <LI><A href="MyMission.asp">已发任务</A> </LI>
   <LI><A href="MyWw.asp">绑定店铺</A> </LI>
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/paipai_task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="Buyhao.asp">绑定买号</A> </LI>
  <LI><A href="MySave.asp">任务仓库</A> </LI></UL>
  </DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG 
src="../images/paipai_task_nav_line.gif"></DIV>
</DIV>
<DIV 
style="MARGIN-BOTTOM: 10px; WIDTH: 910px; PADDING-TOP: 15px; BACKGROUND-COLOR: #ffffff">

<div class="pub_tip4">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td valign="bottom" width="120" align="center"><strong class="f_b_org"><font color="#EC8C11">提示说明</font></strong></td>
        <td><div align="left">1.该页面主要用来绑定和维护用来接任务，购买任务商品的淘宝买号；<a href="http://www.778892.com/news.asp?/1415.html" target="_blank" class="link_o">什么是买号？</a><br />
          2.为了保证您的买号以及发布方网店的安全，请尽量绑定多个备选买号，接手任务时轮换使用；<a href="http://www.778892.com/news.asp?/1480.html" target="_blank" class="link_o">为什么要绑定多个买号？</a><br />
          3.为了保证您的买号以及发布方网店的安全，平台建议每个淘宝买号每日接手任务数不要大于6个；每周不要超过35个!<br />
          4.为了保证刷客网店信誉可信度，系统为每个绑定买号自动设置一个“寿命值”；当该买号买家信誉度达到“寿命值”后将不能接手新的任务；<br /><br /></div></td>
      </tr>
    </table>   
  </div>
<DIV style="BACKGROUND: #ffffff; WIDTH: 900px">
<FORM id=form1 name=form1  action="" onSubmit="return save_onclick12()" method=post>
<input name="action" type="hidden" value="ok">
<DIV></DIV>
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="WIDTH: 910px">
<TABLE width="96%" 
align=center 
style="BORDER-TOP-WIDTH: 0px; MARGIN-TOP: 10px; BORDER-LEFT-WIDTH: 0px; BORDER-BOTTOM-WIDTH: 0px; MARGIN-BOTTOM: 10px; WIDTH: 50%; BORDER-RIGHT-WIDTH: 0px">
  <TBODY>
  <TR>
    <TD align=right height=30>拍拍买入帐号：</TD>
    <TD><div align="left">
      <INPUT id=hao style="WIDTH: 190px" maxLength=20 name=hao>
    （QQ号码）</div></TD></TR>
  <TR>
    <TD height=30><div align="right">操作码：</div></TD>
    <TD><div align="left">
      <input name="czm" type="password" id="czm" size="60" style="WIDTH: 190px">
    </div></TD>
  </TR>
  <TR>
    <TD height=30></TD>
    <TD>&nbsp;</TD>
  </TR>
  <TR>
    <TD height=30></TD>
    <TD><div align="left">
      <INPUT id=Button1 style="WIDTH: 200px; HEIGHT: 30px"  type=submit value=绑定我的拍拍买入帐号 name=Button1>
    </div></TD></TR>
  <TR>
    <TD height=50></TD>
    <TD>&nbsp;</TD>
  </TR></TBODY></TABLE>
</DIV></DIV>
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="WIDTH: 910px">
  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD>
      <TABLE cellSpacing=0 cellPadding=0 align=left border=0>
        <TBODY>
        <TR>
          <TD width=1><IMG height=25 src="../images/mytaobaobj1_3.gif" 
          width=3></TD>
          <TD align=middle width=180 
            background=../images/mytaobaobj1_4.gif><FONT 
            color=#ffffff>已绑定拍拍买入帐号</FONT></TD>
          <TD width=1><IMG height=25 src="../images/mytaobaobj1_6.gif" 
          width=3></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 align=left border=0>
        <TBODY>
       </TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD bgColor=#E9A91E height=4></TD></TR></TBODY></TABLE></DIV></DIV>
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="WIDTH: 910px">
<TABLE width="100%" align="center" cellpadding="5" cellspacing="1" bgcolor="#D0DBE3" >
        <TR bgcolor="#E1F2FB">
          <TD height="30" style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">拍拍买入帐号 </div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">初始买入信誉 </div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">如今已经接手 </div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">该号离黄钻还差多少</div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">状态 </div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 5%"><div align="center">操作 </div></TD>
        </TR>
        <TBODY>
		  <%
	
			
			  	Sql = "select * from "&jieducm&"_xinyu  where username='"&session("useridname")&"' and class=2  order by id desc "
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
			  %>
          <TR>
              <TD bgcolor="#FFFFFF" class=centers><img src="../images/j_z.gif" width="13" height="16"><%=rs("shopname")%></TD>
    <TD bgcolor="#FFFFFF" class=centers><%=rs("num")%></TD>
    <TD bgcolor="#FFFFFF" class=centers><%
	Sqlj = "select count(*) as jxinyu from "&jieducm&"_xinyu where username='"&session("useridname")&"' and classid=2 and shopname='"&rs("num")&"' "
	Set Rsj = Server.CreateObject("Adodb.RecordSet")
	Rsj.Open Sqlj,conn,1,1
	IF not Rsj.Eof Then
	 i=rsj("jxinyu")
	 else
	 i=0
	End IF
	response.Write(i)
	%></TD>
    <TD bgcolor="#FFFFFF" class=centers><%=200-(rs("num")+i)%>
                    <%xiaoyu=(rs("num")+i)
				  if xiaoyu>=200 then
				    sqlr="update "&jieducm&"_xinyu set start=1 where id="&rs("id")&""
                    conn.execute sqlr
				  end if
				  %></TD>
<TD bgcolor="#FFFFFF" class=centers><%if rs("start")="0" then%>可以继续使用 <%else%> 已达到黄钻<%end if%></TD>
    <TD bgcolor="#FFFFFF" class=centers><a href="#1" onClick="javascript:showDialog('你确认要删除此买号吗？',1,7,'?action=del&id=<%=rs("id")%>')" title="你确认要删除此买号吗？">删除</A></TD>
          </TR>
		    <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
        </TBODY>
      </TABLE>
</DIV></DIV>


<DIV> </DIV>

</FORM>
</DIV>

</DIV></div>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
</BODY></HTML>
<SCRIPT language=javascript>
function  save_onclick12()
{

  var Shopkeeper=document.form1.hao.value;
  if(Shopkeeper=="")
  {
    alert("请输入我的刷钻买号！");
	document.form1.hao.focus();
	return false;
	}

  return true;  
}
</SCRIPT>