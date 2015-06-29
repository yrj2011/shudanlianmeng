<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->

<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统**************************************
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
%>
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="../css/center.css" type=text/css rel=stylesheet>
<style type="text/css">
.set {
	clear:both;
}
.set td {
	border-bottom:1px solid #d5d5d5;
	line-height:16px;
	vertical-align:middle;
	padding-left:30px;
}
.set thead td {
	margin-bottom:4px;
	color:#666;
	background:#F8F8F8;
	font-weight:bold;
}
.set tbody tr:hover td {
	background-color:#F8F8F8;
}
.set_bar td {
	background:#F6FBFF url(../images/ico_1.gif) 15px center no-repeat;
	font-weight:bold;
}

.STYLE18 {
	color: #3333FF;
	font-weight: bold;
}
.STYLE23 {color: #3300FF}
.STYLE24 {color: #3333FF}
.STYLE25 {color: #3366CC}
.STYLE26 {color: #3366CC; font-weight: bold; }
.STYLE31 {
	color: #000000;
	font-weight: bold;
}
</style>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   
<!--#include file=../inc/jieducm_top.asp-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<tr>
	    <td>
        <table width="910"  border="0" cellspacing="0" cellpadding="0">
	          <tr>
	          
	            <td width="140" valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"> </td>
    </tr>
</table>



<table id="menuToolBar" width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
    <td valign="top" bgcolor="FFFFFF"><!--#include file="userleft.asp"--></TD>
</TR>
</table></td>
	            <td width="15">&nbsp;</td>
	          
	            <td valign="top">
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"></td>
    </tr>
</table>
                <div class=pp9>
                  <div style="PADDING-BOTTOM: 15px; WIDTH: 97%">
                    <div class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt;官方赏罚名单 &gt;&gt; </div>
                    <div class=pp8>
                      <table width="100%" border="0">
                        <tr>
                          <td width="78%"><strong>官方奖励会员名单</strong></td>
                          <td width="22%"><div align="center"><a href="complist.asp?action=1" class="STYLE11">查看全部</a></div></td>
                        </tr>
                      </table>
                    </div>
                    <table  border="0" cellspacing="1" width="100%" cellpadding="0" class="set">
                      <tr class="title" height="22">
                        <td width="156" height="22" align="center" ><span class="STYLE2"><strong>用户名</strong></span></td>
                        <td width="123" align="center" ><span class="STYLE1"><strong>增值项目</strong></span></td>
                        <td width="123" align="center" ><span class="STYLE1"><strong>奖励数额</strong></span></td>
                        <td width="164" align="center" ><span class="STYLE2"><strong>操作员</strong></span></td>
                        <td width="164" align="center" ><span class="STYLE1"><strong>完成时间</strong></span></td>
                        <td width="323" align="center" ><span class="STYLE1"><strong>奖励<span class="STYLE5">/</span>充值内容</strong></span></td>
                      </tr>
                      <%
sql="SELECT top 10 * FROM "&jieducm&"_chengfa where leibie=1 order by id desc"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Rs.Eof Then
	Response.Write("暂无信息")
Else
Do While Not Rs.Eof
%>
                      <tr class="tdbg" >
                        <td width="156" align="center"><%=rs("username")%></td>
                        <td width="123" align="center"><%if rs("class")=1 then
			 response.Write("存款") 
			elseif rs("class")=2 then
			 response.Write("发布点")
			 elseif rs("class")=3 then
			 response.Write("积分")
		   end if%></td>
                        <td width="123" align="center"><%=rs("num")%></td>
                        <td width="164" align="center"><%=rs("manage")%></td>
                        <td width="164" align="center"><%=rs("now")%> </td>
                        <td width="323" align="center">&nbsp;<%=rs("content")%></td>
                      </tr>
                      <%
Rs.MoveNext
Loop
End IF
%>
                    </table>
                    <div class=pp8>
                      <table width="100%" border="0">
                        <tr>
                          <td width="78%"><strong><span class="STYLE5">官方处罚</span><span class="STYLE5">会员名单</span></strong></td>
                          <td width="22%"><div align="center" class="STYLE5"><a href="complist.asp?action=2"><strong>查看全部</strong></a></div></td>
                        </tr>
                      </table>
                    </div>
                    <table  border="0" cellspacing="1" width="100%" cellpadding="0" class="set">
                      <tr class="title" height="22">
                        <td width="156" height="22" align="center" ><span class="STYLE5"><strong>用户名</strong></span></td>
                        <td width="123" align="center" ><span class="STYLE5"><strong>处罚项目</strong></span></td>
                        <td width="123" align="center" ><span class="STYLE5"><strong>处罚数额 </strong></span></td>
                        <td width="164" align="center" ><span class="STYLE5"><strong>操作员</strong></span></td>
                        <td width="164" align="center" ><span class="STYLE5"><strong>处罚时间</strong></span></td>
                        <td width="323" align="center" ><span class="STYLE5"><strong>处罚/处分原因</strong></span></td>
                      </tr>
                      <%
sql="SELECT top 10 * FROM "&jieducm&"_chengfa where leibie=2 order by id desc"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Rs.Eof Then
	'Response.Write("暂无信息")
Else
Do While Not Rs.Eof
%>
                      <tr class="tdbg" >
                        <td width="156" align="center"><%=rs("username")%></td>
                        <td width="123" align="center"><%if rs("class")=1 then
			 response.Write("存款") 
			elseif rs("class")=2 then
			 response.Write("发布点")
			 elseif rs("class")=3 then
			 response.Write("积分")
		   end if%></td>
                        <td width="123" align="center"><%=rs("num")%></td>
                        <td width="164" align="center"><%=rs("manage")%></td>
                        <td width="164" align="center"><%=rs("now")%> </td>
                        <td width="323" align="center">&nbsp;<%=rs("content")%></td>
                      </tr>
                      <%
Rs.MoveNext
Loop
End IF
%>
                    </table>
                  </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%call closeconn()%>
</BODY></HTML>
