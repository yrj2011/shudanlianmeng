<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
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
<SCRIPT src="../js/Dialog.js" type=text/javascript></SCRIPT>
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
                    <div class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt; 提现列表 &gt;&gt; </div>
                    <div class=pp8><strong>提现列表</strong></div>
                    <br>
                    <br>
					  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
          <tr>
            <td width="125" height="26" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor border-bot"><div align="center">流水号</div></td>
            <td width="142" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor  border-bot"><div align="center">金额</div></td>
            <td width="143" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor  border-bot"><div align="center">提现接收号</div></td>
            <td width="143" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor  border-bot"><div align="center">提现时间 </div></td>
            <td width="94" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor  border-bot"><div align="center">状态</div></td>
            <td width="111" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="border-bot red-bcolor "><div align="center">操作</div></td>
          </tr>
          <%	

   	sql="SELECT * FROM "&jieducm&"_tixian where username='"&session("useridname")&"' order by start desc"
  
			Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then
					
	dim maxperpage,url,PageNo
	 url="ReMoneyList.asp"
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
 DO WHILE NOT rs.EOF AND RowCount>0%>
          <tr>
            <td height="35" align="center" class="border-botdashed"><div align="center"><%=rs("pid")%></div></td>
            <td align="center" class="border-botdashed"><div align="center"><%=rs("ReNum")%>元</div></td>
            <td align="center" class="border-botdashed"> &nbsp;
              <div align="center"><%=rs("ReZfb")%></div></td>
            <td align="center" class="border-botdashed">
              <div align="center"><%=rs("now")%></div></td>
            <td align="center" class="border-botdashed">			<div align="center">
              <%
			if(rs("start")=0) then
			   response.write("<font color=red>处理中</font>")
			elseif (rs("start")=1 )then
			   response.write("提现成功")
			 elseif (rs("start")=2 )then
			   response.write("撤消提现")
			 elseif (rs("start")=3 ) then
			    response.write("已锁定")
			 end if
			%>             
            </div></td>
            <td align="center" class="border-botdashed"><div align="center">
              <%	if(rs("start")=0) then%>
               <a href="#1" onClick="javascript:showDialog('你确认要撤销本此提现吗！',1,7,'tixiandel.asp?id=<%=rs("id")%>')" title="确认要撤销本此提现！"><font color=red>撤销提现</font></a>
               <%
			elseif (rs("start")=1 )then
			   response.write("已支付")
			 elseif (rs("start")=2 )then
			   response.write("已撤销")
			    elseif (rs("start")=3 )then
			   response.write("已锁定")
			 end if
			%> 
            </div></td>
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
<%call closeconn()%>
</BODY></HTML>
