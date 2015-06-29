<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->
<!--#include file="../background.asp"-->
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
    <td valign="top" bgcolor="FFFFFF">
    <!--#include file="userleft.asp"--></TD>
</TR>
</table></td>
	            <td width="15">&nbsp;</td>
	          
	            <td valign="top">
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"></td>
    </tr>
</table>
 <DIV class=pp9>
      <DIV style="PADDING-BOTTOM: 15px; WIDTH: 97%">
      <DIV class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt;官方黑名单 &gt;&gt; </DIV>
      <DIV class=pp8><STRONG>官方黑名单</STRONG></DIV>
      <table width="100%" border="0" cellpadding="0" cellspacing="0"  class="margin6">
          <tr>
            <td width="27%" align="center"><span class="font12h">编号</span></td>
            <td width="21%" height="20" align="center"><span class="font12h">用户名</span>&nbsp;<span class="redcolor">&nbsp;</span></td>
            <td width="29%"><div align="center"><span class="font12h"><strong>操作时间</strong></span></div></td>
            <td width="23%"><div align="center"><span class="font12h">操作人</span></div></td>
            </tr>
				<%	
   	sql="SELECT * FROM  "&jieducm&"_hei  where class='2'  order by id desc"

			Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then
					
	dim maxperpage,url,PageNo
	 url="hei.asp"
	 rs.pagesize=20
	 PageNo=REQUEST("PageNo")
	 if PageNo="" or PageNo=0 then PageNo=1
	 RS.AbsolutePage=PageNo
	 TSum=rs.pagecount
	 maxperpage=rs.pagesize
	 RowCount=rs.PageSize
	   PageNo=PageNo+1
	   PageNo=PageNo-1
	 if CINT(PageNo)>1 then
	    if CINT(PageNo)>CINT(TSum) then
		  response.Write("对不起没有您想要的页数")
          Response.End
	    end if
	 end if
		end if    
     if PageNo<0 then
	    response.Write("没有这一页!")
		Response.End
	 End if
	
	if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
	  %>
       <% DO WHILE NOT rs.EOF AND RowCount>0%>
          <tr>
            <td align="center" ><%=rs("id")%></td>
            <td height="26" align="center" ><%=rs("name")%></td>
            <td align="center" ><%=rs("now")%></td>
            <td align="center" ><%=rs("username")%></td>
            </tr>
			  <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
	   <tr>
            <td height="26" colspan="4" align="center" class="font12h"> <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %></td>
            </tr>
        </table>
		
		
      </DIV>
 </DIV></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%call closeconn()%>
</BODY></HTML>
