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
keeper=Replace(Trim(Request.Form("keeper")),"'","''")
action=Replace(Trim(Request.Form("action")),"'","''")
if action="ok" then
  if Replace(Trim(request("code")),"'","''")<>czm then
  		 Response.Write("<script>alert('操作码错误!');window.location.href='myww.asp';</script>")
		 response.End()
  end if
if not checkmaihao(keeper) then
  Response.Write("<script>alert('你绑定的店铺账号不是有效的账号,请正确输入!');window.location.href='myww.asp';</script>")
  response.End()
end if
		
Sql = "select count(*) as yu  from   "&jieducm&"_keeper where username='"&session("useridname")&"' and class=1"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF not(Rs.Eof) Then
coun=rs("yu")
end if

if jifei<2000 and coun>=2  then
  	Response.Write("<script>alert('2000积分以下，最多可绑定2个店铺!');window.location.href='myww.asp';</script>")
	response.End()
elseif jifei>2000 and jifei<5000 and coun>=4 then
  	Response.Write("<script>alert('2000-5000积分，最多可绑定4个店铺!');window.location.href='myww.asp';</script>")
	response.End()
elseif jifei>5000 and jifei<8000 and coun>=8 then
  	Response.Write("<script>alert('5000-8000积分，最多可绑定8个店铺!');window.location.href='myww.asp';</script>")
	response.End()
elseif jifei>9000  and coun>=10 then
  	Response.Write("<script>alert('9000积分以上，最多可绑定10个店铺!');window.location.href='myww.asp';</script>")
	response.End()
end if

		  
      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_keeper where keeper='"&keeper&"'" ,Conn,3,3  
	  if rs.eof then
	    rs.addnew
		rs("username")=session("useridname")
		rs("keeper")=keeper
		rs("now")=now()
		rs("class")=1
    	rs.update
    	rs.close
		 Response.Write("<script>alert('绑定成功!');window.location.href='myww.asp';</script>")
		 response.End()
	 else
	 	 Response.Write("<script>alert('此淘宝店铺掌柜帐号已被其它用户绑定!');history.back();</script>")
		 response.End()
	 end if
end if
%>	
<HTML><HEAD><title><%=stupname%> -<%=stuptitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<SCRIPT language=javascript>
function  save_onclick12()
{

  var Shopkeeper=document.aspnetForm.keeper.value;
  if(Shopkeeper=="")
  {
    alert("请输入淘宝店铺掌柜帐号！");
	document.aspnetForm.keeper.focus();
	return false;
	}

 var code=document.aspnetForm.code.value;
  if(code=="")
  {
    alert("请输入操作码！");
	document.aspnetForm.code.focus();
	return false;
	}

  return true;  
}
</SCRIPT>
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=../taobao/jieducm_toptao.asp-->
<div align="center">
<DIV 
style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 900px">
<UL id=task_nav>
  <LI><A  href="index.asp">淘宝收藏区</A> </LI>
  <LI><A  href="pushmission.asp">发收藏任务</A> </LI>
  <LI><A  href="ReMission.asp">已接任务</A> </LI>
  <LI><A  href="MyMission.asp">已发任务</A> </LI>
    <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="MyWw.asp">绑定店铺</A> </LI>
  <LI><A href="Buyhao.asp">绑定买号</A> </LI>
  <LI><A href="MySave.asp">任务仓库</A> </LI>
  </UL></DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG 
src="../images/task_nav_line.gif"></DIV>
</DIV>
<DIV 
style="MARGIN-BOTTOM: 10px; WIDTH: 910px; PADDING-TOP: 1px; BACKGROUND-COLOR: #ffffff">
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV 
style="BACKGROUND: url(../images/task_bg_01.jpg) repeat-x 50% top; WIDTH: 910px; LINE-HEIGHT: 38px; HEIGHT: 38px; TEXT-ALIGN: center">我已绑定的淘宝店铺帐号，绑定店铺掌柜号后才能发任务。 
</DIV></DIV>
<FORM id=aspnetForm name=aspnetForm  action=MyWw.asp method=post onSubmit="return save_onclick12()">
<DIV> </DIV>
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="WIDTH: 910px">
<TABLE 
style="BORDER-TOP-WIDTH: 0px; MARGIN-TOP: 10px; BORDER-LEFT-WIDTH: 0px; BORDER-BOTTOM-WIDTH: 0px; MARGIN-BOTTOM: 10px; WIDTH: 70%; BORDER-RIGHT-WIDTH: 0px" 
align=center>
  <TBODY>
  <TR>
    <TD align=right height=30>淘宝店铺掌柜帐号：</TD>
    <TD><div align="left">
      <INPUT id=keeper style="WIDTH: 190px" maxLength=20 name=keeper>
    </div></TD></TR>
  <TR>
    <TD align=right height=30>操作码：</TD>
    <TD><div align="left">
      <INPUT id=code style="WIDTH: 190px"  type=password maxLength=20 name=code>
    </div></TD></TR>
  <TR>
    <TD height=30></TD>
    <TD><div align="left">
      <INPUT id=Button1 style="WIDTH: 200px; HEIGHT: 30px" type=submit value=绑定我的店铺掌柜帐号 name=Button1>
      <input type="hidden" name="action" value="ok">
    </div></TD></TR>
  <TR>
    <TD colSpan=2 height=50><SPAN 
      style="PADDING-LEFT: 120px; COLOR: red">注：一个平台账号最多可以绑十个没被其他账号绑定过的店铺掌柜名，绑定后将无法修改！</SPAN></TD>
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
            color=#ffffff>已绑定淘宝店铺掌柜旺旺帐号</FONT></TD>
          <TD width=1><IMG height=25 src="../images/mytaobaobj1_6.gif" 
          width=3></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 align=left border=0>
        <TBODY>
        <TR>
          <TD>&nbsp;&nbsp;&nbsp; 2000积分以下，可绑定2个店铺；2000-5000积分，可绑定4个店铺；5000-9000积分，可绑定8个店铺；9000积分以上，可绑定10个店铺</TD>
        </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD bgColor=#E9A91E height=4></TD></TR></TBODY></TABLE></DIV></DIV>
<table width="69%" align="center" cellpadding="5" cellspacing="1" bgcolor="#D0DBE3" >
        <tr bgcolor="#E1F2FB">
          <td height="29" class="yell_font" style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">淘宝店铺掌柜 </div></td>
          <td class="yell_font" style="FONT-WEIGHT: bolder; WIDTH: 20%"><div align="center">绑定时间 </div></td>
        </tr>
        <tbody>
		  <%
			  	Sql = "select * from "&jieducm&"_keeper  where username='"&session("useridname")&"' and class=1 order by id desc "
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
			  %>
          <tr>
            <td height="32" bgcolor="#FFFFFF" class=centers><div align="center"><%=rs("keeper")%> </div></td>
            <td bgcolor="#FFFFFF" class=centers><div align="center"><%=rs("now")%> </div></td>
          </tr>
		  	 <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
        </tbody>
      </table>


<DIV></DIV>

</FORM>
</DIV></div>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%rs.close
set rs=nothing
call closeconn()
%>
</BODY></HTML>
