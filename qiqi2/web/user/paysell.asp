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
'Copyright (C) 2008－2009 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
action=request.QueryString("action")
id=request.QueryString("id")
if action="del" then
 Sql= "select * from "&jieducm&"_pay where id="&id&" and jieducm_username='"&session("useridname")&"' and jieducm_start=0"
	 Set rs=server.createobject("ADODB.RECORDSET")
	 Rs.Open Sql,conn,1,1
	 if not rs.eof then
	    car_username=rs("jieducm_username")
	    car_num=rs("jieducm_num")
		car_price=rs("jieducm_price")
	    car_maijia=rs("jieducm_maijia")
		car_start=rs("jieducm_start")
		rs.close
	else
		 Response.Write("<script>alert('无此信息!');history.back();</script>")
        response.End()
	 end if 
	 
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select fabudian From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3 
if not rs.eof then
rs("fabudian")=rs("fabudian")+car_num
rs.update
rs.close
else
 Response.Write("<script>alert('无此信息!');history.back();</script>")
 response.End()
end if

sqlr="delete  from "&jieducm&"_pay where id='"&id&"' and jieducm_username='"&session("useridname")&"' and jieducm_start=0"
conn.execute sqlr
 

	    nums=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=nums
		rs("class")="取消出售发布点"
		rs("content")=session("useridname")&"进行在线取消出售发布点"&car_num&"个"
		rs("price")=0
		rs("jiegou")="取消成功"
    	rs.update
    	rs.close
Response.Write("<script>alert('取消成功！');window.location.href='../user/paysell.asp';</script>")
 response.End()		
end if
%>
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK rel=stylesheet type=text/css href="../css/global.css">
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   
<!--#include file=../inc/jieducm_top.asp-->

        <table width="910"  border="0" align="center" cellpadding="0" cellspacing="0">
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
                
             <TABLE class="user-info nav-table">
  <THEAD>
  <TR>
    <TD>&nbsp;您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt;出售中的发布点 &gt;&gt; </TD>
    <TD width=80 noWrap> </TD>
  </TR></THEAD>
  <TBODY></TBODY></TABLE> 
  <TABLE style="MARGIN-TOP: 5px" class=user-info  >

  <TR>
    <TD class=buy-point colSpan=8><div align="right"><A 
      style="COLOR: red; TEXT-DECORATION: underline" class=renwu-link 
      href="../user/mai.asp">去官方购买</A></div></TD></TR></THEAD>
  <TBODY>
  <TR>
    <TH>买家用户名</TH>
    <TH>价格/数量</TH>
    <TH>指定买家</TH>
    <TH>平均价格 </TH>
    <TH>状态</TH>
    <TH>操作</TH></TR>
  <%	
  action=request.QueryString("action")
if action=1 then
sql="SELECT * FROM "&jieducm&"_pay where jieducm_username='"&session("useridname")&"' and jieducm_start=0 order by id desc"
else
sql="SELECT * FROM "&jieducm&"_pay where jieducm_username='"&session("useridname")&"' and jieducm_start=1 order by id desc"
end if
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
	 url="paysell.asp"
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
 DO WHILE NOT rs.EOF AND RowCount>0

maiuseranme=rs("jieducm_maiuseranme")
Sql2 = "select jifei,vip from "&jieducm&"_register where username='"&maiuseranme&"'"
Set Rs2 = Server.CreateObject("Adodb.RecordSet")
Rs2.Open Sql2,conn,1,1
if not rs2.eof then	
jifei=rs2("jifei")
vip=rs2("vip")
rs2.close
end if				
 %>	
  <TR>
    <TD><div align="center">
	<%if rs("jieducm_start")=0 then%>
	出售中
	<%else%>
	<%=maiuseranme%> 
      <%if vip="1" then %>
      <img src="../images/vip.gif"  alt="尊贵VIP,信誉值：<%call vipxinyu(maiuseranme)%>"  border="0"/>
      <% end if%>
      <%call xinyu(jifei)%> 
	  <%end if%>
    </div></TD>
    <TD><div align="center"><%=formatnumber(rs("jieducm_price"),2,-1)%>/(<%=rs("jieducm_num")%>)个</div></TD>
    <TD><div align="center"><%if rs("jieducm_maijia") ="" then%>无<%else%><%=rs("jieducm_maijia")%><%end if%></div></TD>
    <TD><div align="center"><%=formatnumber((rs("jieducm_price")/rs("jieducm_num")),2,-1)%>元/个</div></TD>
    <TD><div align="center"><%if rs("jieducm_start") =0 then%>出售中<%else%>已出售<%end if%></div></TD>
    <TD class=operate><div align="center">
	<%if rs("jieducm_start") =0 then%>
	<a href="?action=del&id=<%=rs("id")%>" onClick="return confirm('确定要取消操作吗！');">取消</a>
	  <%else%>
	  <IMG  border=0 src="../images/jieducm_xxok.gif">
	  <%end if%>
	   </div></TD>
  </tr>
  <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
	   <TR>
    <TD colspan="6"><div align="center"><% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %></div>      </TD>
    </tr>
        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</BODY></HTML>
