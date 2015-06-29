<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
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
On Error Resume Next 
Server.ScriptTimeOut=9999999 
action=request("action")
if action="ok" then

pingjia=Replace(Trim(Request.Form("pingjia")),"'","''")
url=pingjia
wstr=getHTTPPage(url) 
Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=true
objRegExp.Pattern="<span class=""J_WangWang"" data-nick=""(.*?)""></span>"
set Matches=objRegExp.Execute(wstr)
jieducm_cang=Matches(0).SubMatches(0)
if jieducm_cang="" then
   Response.Write("<script>alert('小号收藏夹地址不正确!');history.back();</script>")
   response.End()
end if
	   
      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_xinyu where shopname='"&jieducm_cang&"' and class=4" ,Conn,3,3  
	  if rs.eof then
	    rs.addnew
		rs("username")=session("useridname")
		rs("shopname")=jieducm_cang
		rs("prourl")=pingjia
		rs("num")=10
		rs("class")=4	
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
<HTML><HEAD><title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<SCRIPT src="../js/Dialog.js" type=text/javascript></SCRIPT>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=../taobao/jieducm_toptao.asp-->
<SCRIPT language=javascript>
function  save_onclick12()
{
 var ProUrl=document.form1.pingjia.value;
  if(ProUrl=="")
  {
    alert("请输入小号收藏夹地址！");
	document.form1.pingjia.focus();
	return false;
	}

  return true;  
}
</SCRIPT>
<div align="center">
<div 
style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
    <div style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
      <div style="FLOAT: left; OVERFLOW: hidden; WIDTH: 900px">
        <ul id=task_nav>
          <li><a  href="index.asp">淘宝收藏区</a> </li>
          <li><a  href="pushmission.asp">发收藏任务</a> </li>
          <li><a href="ReMission.asp">已接任务</a> </li>
          <li><a  href="MyMission.asp">已发任务</a> </li>
		  <LI><A  href="MyWw.asp">绑定店铺</A> </LI>
          <li><a style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="Buyhao.asp">绑定买号</a> </li>
          <li><a href="MySave.asp">任务仓库</a> </li>
        </ul>
      </div>
    </div>
    <div style="CLEAR: both; WIDTH: 910px"><img 
src="../images/task_nav_line.gif"></div>
  </div>
<DIV 
style="MARGIN-BOTTOM: 10px; WIDTH: 910px; PADDING-TOP: 15px; BACKGROUND-COLOR: #ffffff">

 <div class="pub_tip4">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td valign="bottom" width="120" align="center">&nbsp;</td>
        <td><div align="left">注意：<br>
          ①在淘宝收藏时，务必勾选“<STRONG id="GongKaiFavSpan" title="<img src='images/GongKaiFav.gif' alt='公开收藏要打√哦' />" jQuery1292576186015="8">公开收藏</STRONG>”，否则平台检测不到您是否收藏了此商品(或店铺)。 <br>
          <BR>
          ②成人用品店铺(或商品)不能直接勾选“<STRONG>公开收藏</STRONG>”，可先收藏，收藏后去收藏夹找到该条收藏记录，点“编辑”，勾选“<STRONG>公开收藏</STRONG>”即可。 </div></td>
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
    <TD height=30><div align="right">小号收藏夹地址：</div></TD>
    <TD><div align="left">
      <input name="pingjia" type="text" id="pingjia" size="60" style="WIDTH: 190px">
    </div></TD>
  </TR>
  <TR>
    <TD height=30></TD>
    <TD><div align="left"><a href="../news.asp?/1397.html" target="_blank">收藏夹地址从哪找</a></div></TD>
  </TR>
  <TR>
    <TD height=30></TD>
    <TD><div align="left">
      <INPUT id=Button1 style="WIDTH: 200px; HEIGHT: 30px"  type=submit value=绑定我的淘宝收藏帐号 name=Button1>
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
            color=#ffffff>已绑定的淘宝收藏小号</FONT></TD>
          <TD width=1><IMG height=25 src="../images/mytaobaobj1_6.gif" 
          width=3></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 align=left border=0>
        <TBODY>
        <TR>
          <TD>&nbsp;</TD>
        </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD bgColor=#E9A91E height=4></TD></TR></TBODY></TABLE></DIV></DIV>
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="WIDTH: 910px">
<TABLE width="100%" align="center" cellpadding="5" cellspacing="1" bgcolor="#D0DBE3" >
        <TR bgcolor="#E1F2FB">
          <TD height="30" style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">淘宝买入帐号 </div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">收藏夹地址</div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">添加时间</div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 5%"><div align="center">操作 </div></TD>
        </TR>
        <TBODY>
		  <%
			 
			
			  	Sql = "select * from "&jieducm&"_xinyu  where username='"&session("useridname")&"' and class=4  order by id desc "
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
			  %>
          <TR>
              <TD bgcolor="#FFFFFF" class=centers><img src="../images/j_z.gif" width="13" height="16"><%=rs("shopname")%></TD>
    <TD bgcolor="#FFFFFF" class=centers><label>
      <input name="textfield" type="text" size="50" value="<%=rs("prourl")%>">
    </label></TD>
    <TD bgcolor="#FFFFFF" class=centers><%=rs("now")%></TD>
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
<%rs.close
set rs=nothing
call closeconn()
%>
</BODY></HTML>
