<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
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
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>用 户 管 理</title>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>

<body>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr>
    <td colspan="2" align="center" class="title"><strong>用 户 管 理</strong></td>
  </tr>
  <tr class="tdbg">
    <td width="70" height="30" ><strong>管理导航：</strong></td>
    <td><a href="usermanage.asp">所有会员 </a>|<a href="usermanage.asp?action=pu">未审核会员 </a>| <a href="usermanage.asp?action=vip">已审核会员</a>| <a href="usermanage.asp?action=taobao">淘宝区会员 </a>|<a href="usermanage.asp?action=paipai">拍拍区会员 </a>| <a href="usermanage.asp?action=youa">有啊区会员 </a>|<a href="usermanage.asp?action=you">有余额 |</a> <a href="usermanage.asp?action=wu">无余额</a> | <a href="?action=isjie">禁止接任务</a> | <a href="?action=isfa">禁止发任务</a>| <a href="?action=isdun">禁止兑换</a> | <a href="?action=islogin">禁止登录</a> | <a href="?action=isgive">禁止赠送</a> | <a href="?action=isgives">赠送不收手续费</a> | <a href="?action=ismessage">禁止站内信</a>| <a href="?action=wangwang">已绑定手机</a> | <a href="?action=vipuser">VIP会员</a> | <a href="?action=lastnow">最新登陆</a>|<a href="?action=zuan">登陆最多</a></td>
  </tr>
</table>
<table border=0 width=100% align=center bgcolor="#FFFFFF">
  <tr>
            <td width="100%"> 
              <div align="center">
                <form name="form1" method="post" action="usermanage.asp?action=sear">

                  <IMG src="../../images/search.gif" width="16" height="16" align=absMiddle>搜索用户名：
                  <input class=input1 size=25 name=text>
                  
                  <input type="checkbox" name="hao" id="hao" value="1">
                  绑定淘宝小号搜索　 
                  <input type="checkbox" name="da" id="da" value="1">查询绑定淘宝大号
				   <input type="checkbox" name="phone" id="phone" value="1">
				     查询绑定手机号
				     <input type="checkbox" name="qq" id="qq" value="1">查询qq号
<input name="submit" type=submit class=input2 value=" 搜 索 ">
                </form> 
            </div></td>
  </tr>
          
          </td>
</table>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center><tr></tr>
</table>
<table width='100%' border="0" cellpadding="0" cellspacing="0" ><tr>
   
      <table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
        <tr>
          <td colspan="2" align="left" class="title"></td>
        </tr>
        <tr valign="middle" class="tdbg">
          <td height="22"></td>
          <td width="200" height="22" align="right"></td>
        </tr>
      </table>
     <td><table class="border" border="0" cellspacing="1" width="100%" cellpadding="0">
          <tr class="title" height="22">
            <td width="131" height="22" align="center" ><strong>用户名</strong></td>
            <td width="115" align="center" ><strong>联系</strong></td>
            <td width="297" align="center" ><strong>存款 发布点  积分   注册时间</strong></td>
            <td width="159" align="center" ><strong>发任务数</strong></td>
            <td width="157" align="center" ><strong>接任务数</strong></td>
            <td width="270" align="center" ><strong>组操作</strong></td>
            <td width="248" align="center" ><strong>操作</strong></td>
            <!--<td width="40" align="center" ><strong>已审核</strong></td>-->
          </tr>
       
			<%	
			action=request("action")
			hao=request("hao")
			da=request("da")
			phone=request("phone")
			qq=Replace(Trim(Request.Form("qq")),"'","''")
			if action="you" then
			sql="SELECT * FROM "&jieducm&"_register where cunkuan<>0  order by cunkuan desc"
			elseif action="wu" then
			sql="SELECT * FROM "&jieducm&"_register where cunkuan=0 order by id desc"
			elseif action="sear"  and hao<>"1"  and da<>"1" and phone <>"1"  and qq<>"1" then
			   	sql="SELECT * FROM "&jieducm&"_register where username like '%"&request("text")&"%'  order by id desc"
			elseif action="sear" and qq="1" then	
			   sql="SELECT * FROM "&jieducm&"_register where qq='"&request("text")&"' order by id desc"	
			elseif  action="sear" and phone="1" then
			   Sql = "select * from "&jieducm&"_register where dst="&request("text")&""	
			elseif action="pu" then
			   sql="SELECT * FROM "&jieducm&"_register where level1='0' order by id desc"
			elseif action="vip" then
				sql="SELECT * FROM "&jieducm&"_register where level1='1' order by id desc"
			elseif action="isjie" then
				sql="SELECT * FROM "&jieducm&"_register where isjie=1 order by id desc"
			elseif action="isfa" then
				sql="SELECT * FROM "&jieducm&"_register where isfa=1 order by id desc"
			elseif action="isdun" then
				sql="SELECT * FROM "&jieducm&"_register where isdun=1 order by id desc"
			elseif action="islogin" then
				sql="SELECT * FROM "&jieducm&"_register where islogin=1 order by id desc"
			elseif action="isgive" then
				sql="SELECT * FROM "&jieducm&"_register where isgive=1 order by id desc"
			elseif action="isgives" then
				sql="SELECT * FROM "&jieducm&"_register where isgives=1 order by id desc"
			elseif action="ismessage" then
				sql="SELECT * FROM "&jieducm&"_register where ismessage=1 order by id desc"
			elseif action="wangwang" then
				sql="SELECT * FROM "&jieducm&"_register where dst<>0 order by id desc"
			elseif action="taobao" then
				sql="SELECT * FROM "&jieducm&"_register where taobao=1 order by id desc"
			elseif action="paipai" then
				sql="SELECT * FROM "&jieducm&"_register where paipai=1 order by id desc"
			elseif action="youa" then
				sql="SELECT * FROM "&jieducm&"_register where youa=1 order by id desc"	
			elseif action="vipuser" then
				sql="SELECT * FROM "&jieducm&"_register where vip=1 order by id desc"		
			elseif action="lastnow" then
				sql="SELECT * FROM "&jieducm&"_register  order by lastnow desc"	
			elseif action="zuan" then
					sql="SELECT * FROM "&jieducm&"_register  order by zuan desc"					
			elseif action="sear" and hao="1" then
			    Sql = "select * from "&jieducm&"_register where username=(Select username From "&jieducm&"_xinyu where shopname='"&request("text")&"' and class=1) order by id desc"
		    elseif  action="sear" and da="1" then
			   Sql = "select * from "&jieducm&"_register where username=(Select username From "&jieducm&"_keeper where keeper='"&request("text")&"')  order by id desc"
			else
			   	sql="SELECT * FROM "&jieducm&"_register  order by level1 ,id desc"
			end if
			Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then
	dim maxperpage,url,PageNo
	if action="" then
	 url="usermanage.asp"
	 elseif action="sear"  then
	url="usermanage.asp?action=sear&text="&request("text")&""
	else
	 url="usermanage.asp?action="&action&""
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
	i=0
	if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
	 DO WHILE NOT rs.EOF AND RowCount>0
	 username=rs("username")
	 qq=rs("qq")
	 registerip=rs("registerip")
	 LastLoginIP=rs("LastLoginIP")
	 %>
        <form name="myform<%=i%>" method="Post" action="useredit2.asp" >
		  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;">
		    <td width="131" align="center">
             
            <a href="admin_recordu.asp?username=<%=username%>" title="查看此会员的操作记录"> <%=username%></a><%if Replace(Trim(rs("vip")),"'","''")	="1" then%><img src="../images/jieducm_vip.gif"  title="vip会员" /><%end if%><br>推荐人：<%=rs("tjr")%><br>
			<%if rs("level1")="0" then %><font color="#FF0000"><strong>未审核</strong></font><%else%><font color="#0000FF"><strong>已审核</strong></font><%end if%><br>
			<%if rs("activestart")="0" then %><font color="#FF0000"><strong>未邮件激活</strong></font>
			<%elseif  rs("activestart")="1" then%><font color="#0000FF"><strong>已邮件激活</strong></font><%end if%>
			<br>
				<%if rs("dst")="0" then %><font color="#FF0000"><strong>未绑定手机</strong></font>
			<%else%><font color="#0000FF"><strong>已绑定手机<%=rs("dst")%></strong></font><%end if%>
			<br><%if rs("tribeid")=0 then%><font color="#FF0000"><strong>未加入部落</strong></font>
			<%else%><font color="#0000FF"><strong>已加入部落</strong></font><%end if%>
			</td>
            <td width="115" align="center">
			<%Sql1 = "select  count(*) as jieducm_tjr  from "&jieducm&"_register where tjr='"&username&"' "
            Set Rs1 = Server.CreateObject("Adodb.RecordSet")
				Rs1.Open Sql1,conn,1,1
				if NOT rs1.EOF  then
				jieducm_tjr=rs1("jieducm_tjr")
            end if%>推荐人数：<font color="#FF0000"><strong><%=jieducm_tjr%></strong></font>人
			 <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=qq%>&site=qq&menu=yes">
			<img src="http://wpa.qq.com/pa?p=1:<%=qq%>:41"  border="0"  title="QQ：<%=qq%>" /></a></td>
            <td width="297" align="center"><div align="left"><a href="chongf.asp?id=<%=rs("id")%>" title="对此会员后台充值"><font color="#FF0000"><strong>存款：
                <%
if rs("cunkuan")=0 then
response.Write("0.0")
else
response.Write(formatnumber(rs("cunkuan"),2))
end if
%> 
                发布点：
                <%
if rs("fabudian")=0 then
response.Write("0.0")
else
response.Write(formatnumber(rs("fabudian"),2))
end if
%>
  积分：<%=rs("jifei")%></strong></font></a><br>
  注册时间：<%=rs("now")%><br>
  注册ip：<a href="http://www.ip138.com/ips.asp?ip=<%=registerip%>&action=2" target="_blank" title="查看ip所在地"><%=registerip%></a>
  &nbsp;<a href="jieducm_ip.asp?ip=<%=registerip%>"><font color="#FF0000">锁定此IP</font></a><br>
  最后登陆时间：<%if isnull(RS("lastnow")) then%><%=rs("now")%><%else%><%=rs("lastnow")%><%end if%><BR> 
  最后登陆ip：
  <%if not isnull(rs("LastLoginIP")) then%>
  <a href="http://www.ip138.com/ips.asp?ip=<%=LastLoginIP%>&action=2" target="_blank" title="查看ip所在地"><%=LastLoginIP%></a>
  &nbsp;<a href="jieducm_ip.asp?ip=<%=LastLoginIP%>"><font color="#FF0000">锁定此IP</font></a>
  <%else%>
  <a href="http://www.ip138.com/ips.asp?ip=<%=registerip%>&action=2" target="_blank" title="查看ip所在地"><%=registerip%></a>
  &nbsp;<a href="jieducm_ip.asp?ip=<%=registerip%>"><font color="#FF0000">锁定此IP</font></a>
  <%end if%>
  
  <br>
  登陆次数：<%=rs("zuan")%>次
  <br>
  <a href="jieducm_usermoney.asp?username=<%=username%>"><Font color="#FF0000"><strong>查看此会员资金数据</strong></Font></a>
  </div></td>        
            <td width="159" align="center"> 
           <% Sql1 = "select  count(*) as renwu  from "&jieducm&"_zhongxin where username='"&username&"' "
            Set Rs1 = Server.CreateObject("Adodb.RecordSet")
				Rs1.Open Sql1,conn,1,1
				if NOT rs1.EOF  then
				renwu=rs1("renwu")
            end if
			
			totalprices=0
			Sql1 = "select  sum(price) as totalprices  from "&jieducm&"_zhongxin where username='"&username&"'  and classid<>'6'"
            Set Rs1 = Server.CreateObject("Adodb.RecordSet")
			Rs1.Open Sql1,conn,1,1
			if NOT rs1.EOF  then
				totalprices=rs1("totalprices")
		    else
				totalprices=0
            end if
		    rs1.close
			  %>  
			  <%=renwu%>个<br>
			  
			  <a href="admin_mymission.asp?sort=ok&text=<%=username%>">查看已发任务</a><br><br>
			  已发任务总额：<font color="#FF0000"><%if isnull(totalprices) then%> 0 <%else%><%=totalprices%><%end if%></font>元
            </td> 
            <td width="157" align="center">     
		  <% Sql2 = "select  count(*) as jiewu  from "&jieducm&"_jieshou where username='"&username&"' "
            Set Rs1 = Server.CreateObject("Adodb.RecordSet")
				Rs1.Open Sql2,conn,1,1
				if NOT rs1.EOF  then
				jiewu=rs1("jiewu")
				end if
				
			Sql1 = "select  sum(price) as totalpricesj  from "&jieducm&"_jieshou where username='"&username&"'  and classid<>'6'"
            Set Rs1 = Server.CreateObject("Adodb.RecordSet")
			Rs1.Open Sql1,conn,1,1
			if NOT rs1.EOF  then
				totalpricesj=rs1("totalpricesj")
				else
				totalpricesj=0
            end if
				rs1.close  %> 
				  <%=jiewu%>个<br>
				    <a href="admin_MyMissionjie.asp?sort=ok&text=<%=username%>">查看已接任务</a><br><br>
				  已接任务总额：<font color="#FF0000"><%if isnull(totalpricesj) then%> 0 <%else%><%=totalpricesj%><%end if%></font>元
		    </td>         
            <td width="270" align="center">
               <label> <input type="checkbox" name="level1" id="level1"  value="1" <%if rs("level1")="1" then%> checked <% end if%> > 审核会员 </label>
              <label> <input type="checkbox" name="taobao" id="taobao"  value="1" <%if rs("taobao")="1" then%> checked <% end if%> > 淘宝区会员<br>
</label> <label> <input type="checkbox" name="paipai" id="paipai"  value="1" <%if rs("paipai")="1" then%> checked <% end if%> > 拍拍区会员 </label>

               <label> <input type="checkbox" name="isjie" id="isjie"  value="1" <%if rs("isjie")=1 then%> checked <% end if%> > 禁止接任务 </label><br>
               <label><input type="checkbox" name="isfa" id="isfa" value="1" <% if rs("isfa")=1 then%> checked <% end if%>>禁止发任务</label>
            <label> <input type="checkbox" name="isdun" id="isdun" value="1" <%if rs("isdun")=1 then%> checked <%end if%> >禁止兑换</label><br>
             <label><input type="checkbox" name="islogin" id="islogin" value="1" <%if rs("islogin")=1 then%> checked <%end if%>>禁止登录</label>
             
           
              <label> <input type="checkbox" name="isgive" id="isgive" value="1" <%if rs("isgive")=1 then%> checked <% end if%>>禁止赠送</label><br>
             <label><input type="checkbox" name="isgives" id="isgives" value="1" <% if rs("isgives")=1 then%> checked <% end if%>>赠送不收手继费</label>
              <label><input name="ismessage" type="checkbox" id="ismessage" value="1" <% if rs("ismessage")=1 then%> checked <% end if%>>禁止站内信</label>
 
              <input type="hidden" name="id" id="id" value="<%=rs("id")%>">
              <input type="hidden" name="username1" id="username1" value="<%=rs("username")%>">
            <br>            </td>
          
            <td width="248" align="center" >
			<a href="userunionlist.asp?username=<%=rs("username")%>">查看此会员的推广记录</a><br>
			<a onClick="javascript:document.forms['myform<%=i%>'].submit();" style="cursor:pointer;">修改组操作</a>|<a href="chengfa.asp?username=<%=username%>">奖励惩罚</a>
              <br>
             <a href="useredit.asp?username=<%=username%>"> 修改详情</a>| <a onClick="return confirm('确定要删除此会员吗？此操作不能恢复！');" href="userdel.asp?id=<%=rs("id")%>">删除</a><br>
            
          <a href="userbuyhao2.asp?id=<%=rs("id")%>">绑定买号</a>|<a href="userbuyhao.asp?username=<%=username%>">管理买号</a>
		  <br>
		  <a href="usermyww.asp?id=<%=rs("id")%>">绑定掌柜</a>| <a href="userbuyww.asp?username=<%=username%>">管理掌柜</a><br>
		  <%if rs("codenum")<>0 then%><a href="uphonelook.asp?id=<%=rs("id")%>">查看手机验证码</a><br><%end if%>
		   <a href="jieducm_vipuser.asp?username=<%=username%>">设置为vip会员</a>
		  </td>   
          </tr></form>
 <%
	i=i+1
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
          </table>
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="border">
  <tr class="tdbg">
    <td height="30" colspan="2">
&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
  <tr bgcolor="#FFFFFF"><td  ><div align="center"><br>
    <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %>
  </div></td></tr>
</table>
</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" align="center" background=images/admin_bg_bottom.gif><font color=#cccccc>Copyright &copy; 2013</font> <a href='http://www.778892.com' target="_blank"><font color=#ff6600><strong>七七传媒-七七互刷平台</strong></font></a> <font color=#cccccc>. All Rights Reserved.</font></td>
  </tr>
</table>

<%rs.close
set rs=nothing
set rs1=nothing
conn.close
set conn=nothing%>
</body>
</html>
