<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->
<!--#include file="../md5.asp"-->
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
From_url = Cstr(Request.ServerVariables("HTTP_REFERER"))
Serv_url = Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(From_url,8,len(Serv_url)) <> Serv_url or Request.ServerVariables("HTTP_REFERER")="" then
response.write "请求超时,请关闭本页面并重新登陆即可!"
response.end
end If
action=request.form("action")
if action="1" then
num=Replace(Trim(Request.Form("num")),"'","''")
price=Replace(Trim(Request.Form("price")),"'","''")
username=Replace(Trim(Request.Form("username")),"'","''")

if int(fabudian)<int(stuppaynum) then
 	Response.Write("<script>alert('发布点低于"&stuppaynum&"不能出售!');history.back();</script>")
	response.End()	
elseif(int(fabudian)<int(num)) then
    Response.Write("<script>alert('你出售的发布点不能大于你现有的发布点!');history.back();</script>")
	response.End()	
elseif czm<>request.form("czm") then
	Response.Write("<script>alert('操作码不正确!');history.back();;</script>")
	response.End()	
elseif stuppayis=0 then
    Response.Write("<script>alert('此功能已关闭！');history.back();</script>")
   response.End()
elseif jifei< stuppayjifei	then
    Response.Write("<script>alert('积分高于"&stuppayjifei&"才可以使用此功能！');history.back();</script>")
   response.End()
elseif int(num)<1 then
	Response.Write("<script>alert('输入有误请重新输入!');history.back();;</script>")
	response.End() 
elseif not isNumeric(num) then
   Response.Write("<script>alert('输入有误请重新输入！');history.back();</script>")
   response.End()
elseif not isNumeric(price) then
   Response.Write("<script>alert('输入有误请重新输入！');history.back();</script>")
   response.End()
elseif int(price)<1 then
   Response.Write("<script>alert('输入有误请重新输入！');history.back();</script>")
   response.End()
end if

if username<>"" then
     Set rs=server.createobject("ADODB.RECORDSET")
	 rs.open "Select * From "&jieducm&"_register where username='"&username&"'" ,Conn,1,1  
	   if rs.eof then
	       Response.Write("<script>alert('指定买家用户名不存在');history.back();</script>")
		   response.End()
       elseif session("useridname")=username then
	      Response.Write("<script>alert('指定买家用户名不能是自己!');history.back();</script>")
		  response.End()
	  end if
end if	  

if vip=1 then
jieducm_shou=(stupvippaynum/100)*num
else
jieducm_shou=(stuppupaynum/100)*num
end if

if (num+jieducm_shou)>fabudian then
    Response.Write("<script>alert('你出售的发布点不能大于你现有的发布点!');history.back();</script>")
	response.End()	
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select fabudian From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
if rs.eof then
 Response.Write("<script>alert('无此会员');window.location.href='pay.asp';</script>")
response.End()
else
rs("fabudian")=rs("fabudian")-num-jieducm_shou
rs.update
rs.close
end if	  

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_pay " ,Conn,3,3  
rs.addnew
rs("jieducm_username")=session("useridname")
rs("jieducm_num")=num
rs("jieducm_price")=price
rs("jieducm_maijia")=username
rs("jieducm_datatime")=now()
rs("jieducm_start")=0
rs.update
rs.close
	    nums=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=nums
		rs("class")="出售发布点"
		rs("content")=session("useridname")&"进行在线出售发布点"&num&"个,设置出售总价格为"&price&"元扣手续费"&jieducm_shou&"个发布点"
		rs("price")=0
		rs("jiegou")="出售成功"
    	rs.update
    	rs.close  
		
 Response.Write("<script>alert('发布成功！');window.location.href='../user/paypoint.asp';</script>")
 response.End()	
end if
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
 <SCRIPT language=javascript>
function  save_onclick()
{
    var num=document.form.num.value;
  if(num=="")
  {
    alert("请输入要出售的发布点数量！");
	document.form.num.focus();
	return false;
	}
 if(!Number(num))
	  
  {
    alert("请您输入数字!");
	document.form.num.focus();
	return false;
	}
if ((document.form.num.value.indexOf("-")   ==   0)||!(document.form.num.value.indexOf(".")   ==   -1)){   
  alert("出售数量不能为小数或负数");   
  document.form.num.focus();   
  return   false;   
  }
 
  var price=document.form.price.value;
  if(price=="")
  {
    alert("请输入出售的总价格！");
	document.form.price.focus();
	return false;
	}
	  
 var czm=document.form.czm.value;
  if(czm=="")
  {
    alert("请输入您的操作码！");
	document.form.czm.focus();
	return false;
	}
  
  return true;  
}
</script>   
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
                    <div class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt;出售发布点 &gt;&gt; </div>
                    <div class=pp8><strong>出售发布点</strong></div>
                    <div style="MARGIN-TOP: 10px; LINE-HEIGHT: 200%">
                      <ul>
                        <li>
<SPAN style="COLOR: red">*所扣取的手续费为"摊租",取消出售不返还 手续费,请定位好合适的价格再出售!如果您不想出售,您还可以通过官方兑换直接换取现金!</SPAN></li>
                      </ul>
                    </div><br> 
                    <form action="" method="POST" name="form" id="form" onSubmit="return save_onclick()">
                      <table width="679" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="35" align="right"><div align="right"></div></td>
                          <td> 您当前可用发布点是：<%=fabudian%> 个 </td>
                        </tr>
                        <tr>
                          <td width="123" height="35" align="right"><div align="right">要出售的发布点数量：</div></td>
                          <td width="556"><input  name=num id=num  onKeyUp="if(isNaN(value))execCommand('undo')" maxlength="4">
                            个 拥有<%=stuppaynum%>个发布点才能出售!&nbsp; vip会员扣<%=stupvippaynum%>%手续费。普通会员扣<%=stuppupaynum%>%手续费</td>
                        </tr>
                        <tr>
                          <td height="35"><div align="right">设置出售的总价格：</div></td>
                          <td><input type="text" name="price" id="price"  onKeyUp="if(isNaN(value))execCommand('undo')">
                          出售价格越低购买的人越多</td>
                        </tr>
                        <tr>
                          <td height="35"><div align="right">指定买家用户名：</div></td>
                          <td><label>
                            <input name="username" type="text" id="username">
                          你可发指定仅某一个用户才可以购买这些发布点</label></td>
                        </tr>
                        <tr>
                          <td height="35"><div align="right">操作码：</div></td>
                          <td><label>
                            <input name="czm" type="password" id="czm">
                          </label></td>
                        </tr>
                        <tr>
                          <td height="35">&nbsp;</td>
                          <td><input type="submit" name="button" id="button" value="确认出售发布点" <%if stuppayis=0 or jifei<stuppayjifei then%> disabled <%end if%> />
                              <input type="hidden" name="action" value="1">
                              &nbsp;&nbsp;  &nbsp; 积分高于<%=stuppayjifei%>点才可以使用</td>
                        </tr>
                      </table>
                    </form>
            
					
				  </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close()
set rs=nothing
call closeconn()%>
</BODY></HTML>
