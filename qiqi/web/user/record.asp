<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
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
                <div class=pp9>
                  <div style="PADDING-BOTTOM: 15px; WIDTH: 97%">
                    <div class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt; 操作记录 &gt;&gt; </div>
                    <div class=pp8><strong>*<a href="Record.asp">所有操作记录列表</a> *<a href="Record.asp?action=ren">任务操作列表</a> * <a href="Record.asp?action=chong">充值列表</a> * <a href="Record.asp?action=chongmanage">后台充值列表</a> * <a href="Record.asp?action=zeng">增值操作列表</a> * <a href="Record.asp?action=ti">提现操作列表</a> * <a href="Record.asp?action=manage">管理操作列表</a></strong></div>
                   
                  </DIV>
                  <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" >
          
          <tr>
            <td width="125" height="26" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor border-bot">流水号</td>
            <td width="142" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor  border-bot">类型</td>
            <td width="143" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor  border-bot">详细</td>
            <td width="80" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor  border-bot">金额</td>
            <td width="94" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor  border-bot">结果</td>
            <td width="111" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="border-bot red-bcolor ">时间</td>
          </tr>
          <%	
		  action=request("action")
		  if action="ren" then
            sql="select * from "&jieducm&"_record where username='"&session("useridname")&"' and (class='淘宝区任务' or class='VIP区任务' or class='拍拍区任务' or class='有啊区任务') order by id desc"
		  elseif action="chong" then
             sql="SELECT * FROM "&jieducm&"_record where   username='"&session("useridname")&"' and (class='充值卡充值' or class='网银充值' or class='支付宝充值') order by id desc"
		  elseif action="chongmanage" then
             sql="SELECT * FROM "&jieducm&"_record where   username='"&session("useridname")&"' and (class='任务' or class='充值' or class='增值' or class='提现' or class='其它') order by id desc"
		 elseif action="zeng" then
sql="select * from "&jieducm&"_record where username='"&session("useridname")&"' and (class='发布点赠送' or class='购买发布点'  or class='积分换发布点'  or class='发布点换存款')  order by id desc"
         elseif action="ti" then
            sql="select * from "&jieducm&"_record where username='"&session("useridname")&"' and class='申请提现'  order by id desc"
		elseif action="manage" then
             sql="select * from "&jieducm&"_record where username='"&session("useridname")&"' and  (class='清理买家' or class='清理删除' or class='返回上级' or class='确认发货' or class='买方好评' or class='卖方完成' or class='惩罚记录' or class='申请VIP区用户' or class='申请退出VIP区') order by id desc"
		  else  
      	        sql="SELECT * FROM "&jieducm&"_record where username='"&session("useridname")&"' order by id desc"
         end if
			Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then
					
	dim maxperpage,url,PageNo
	if action="" then
	   url="record.asp"
	else
	   url="record.asp?action="&action&""
  end if
	 rs.pagesize=10
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
 DO WHILE NOT rs.EOF AND RowCount>0%>
          <tr>
            <td height="35" align="center" class="border-botdashed"><%=rs("num")%></td>
            <td align="center" class="border-botdashed"><%=rs("class")%></td>
            <td align="center" class="border-botdashed"><%=rs("content")%></td>
            <td align="center" class="border-botdashed"> &nbsp;<%=rs("price")%></td>
            <td align="center" class="border-botdashed">&nbsp; <%=rs("jiegou")%></td>
            <td align="center" class="border-botdashed"><%=rs("now")%> </td>
          </tr>
             <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
          <tr>
            <td height="35" colspan="6" align="center" class="border-botdashed"><% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %></td>
            </tr>
        </table>
                  </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
call closeconn()%>
</BODY></HTML>
