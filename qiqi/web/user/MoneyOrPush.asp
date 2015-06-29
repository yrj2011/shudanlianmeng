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
'Copyright (C) 2008－2011 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------

if stupphonestart=1 and dst =0 then
	Response.Write("<script>alert('请先绑定手机才能进行其它操作!');window.location.href='../user/Center_Userlock.asp';</script>")
	response.End()
end if
				
				
action=HtmlEncode(trim(request.form("action")))
if action="1" then
  geshu=(trim(request.form("ToUser")))
if stupfhc="0" then
 	 Response.Write("<script>alert('发布点兑换现金以被管理员关闭，\n如需兑换，请联系客服!');history.back();</script>")
	 response.End()
elseif  isdun=1 then
	Response.Write("<script>alert('你已被管理员禁止了兑换功能!');history.back();</script>")
	response.End()
elseif fabudian<stupfhcshu then
 	Response.Write("<script>alert('发布点低于"&stupfhcshu&"不能兑换!');history.back();</script>")
	response.End()		
elseif(int(fabudian)<int(geshu)) then
    Response.Write("<script>alert('你发布点不够,不能兑换!');history.back();</script>")
	response.End()	
elseif czm<>request("czm") then
	Response.Write("<script>alert('操作码不正确!');history.back();;</script>")
	response.End()
elseif int(geshu)<1 then
	Response.Write("<script>alert('输入有误!');history.back();;</script>")
	response.End() 
elseif not isNumeric(geshu) then
    Response.Write("<script>alert('输入有误请重新输入！');history.back();</script>")
    response.End()
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
if rs.eof then
 Response.Write("<script>alert('无此会员');window.location.href='moneyorpush.asp';</script>")
response.End()
else
rs("fabudian")=rs("fabudian")-geshu
rs("cunkuan")=rs("cunkuan")+(geshu*stupfa)
rs.update
rs.close
 end if	  

	  num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="发布点换存款"
		rs("content")=session("useridname")&"用"&geshu&"个发布点换存款"
		rs("price")=geshu*stupfa
		rs("jiegou")="发布点换存款成功"
    	rs.update
    	rs.close		
		call check_jieducm_name(session("useridname"))   			
    Response.Write("<script>alert('兑换成功!');window.location.href='moneyorpush.asp';</script>")
	response.End()	
    set rs=nothing
elseif action="3" then
     give1=HtmlEncode(trim(request.form("givenum")))
	 touser=HtmlEncode(trim(request.form("touser")))
	 give=give1+give1*(stupshouxu/100)

	if isgive=1 then
		 Response.Write("<script>alert('你已被管理员禁止了赠送功能!');history.back();;</script>")
	     response.End()
	elseif czm<>request("czm") then
		Response.Write("<script>alert('操作码不正确!');history.back();;</script>")
	     response.End()
	elseif((fabudian)<=(give)) then
		 Response.Write("<script>alert('你赠送的发布点不能大于你的发布点!');history.back();</script>")
		response.End() 
	elseif int(give1)<1 then
	   Response.Write("<script>alert('输入有误!');history.back();</script>")
		response.End() 
	elseif not isNumeric(give1) then
          Response.Write("<script>alert('输入有误请重新输入！');history.back();</script>")
          response.End()
	 end if	
	 
	 Set rs=server.createobject("ADODB.RECORDSET")
	 rs.open "Select * From "&jieducm&"_register where username='"&touser&"'" ,Conn,3,3  
	   if rs.eof then
	       Response.Write("<script>alert('无此会员');history.back();</script>")
		   response.End()
       elseif session("useridname")=touser then
	      Response.Write("<script>alert('不能赠送给自己发布点!');history.back();</script>")
		  response.End()
	   else
		rs("fabudian")=rs("fabudian")+give1
    	rs.update
    	rs.close
	  end if
	  
if isgives=1 then
sqlr="update "&jieducm&"_register set fabudian=fabudian-"&give1&"  where username='"&session("useridname")&"'"
conn.execute sqlr
else	  
sqlr="update "&jieducm&"_register set fabudian=fabudian-"&give&"  where username='"&session("useridname")&"'"
conn.execute sqlr
end if		
		
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="发布点赠送"
		rs("content")=session("useridname")&"赠送给"&touser&"会员"&give1&"个发布点"
		rs("price")=0
		rs("jiegou")="发布点赠送成功"
    	rs.update
    	rs.close
 				
	Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=touser
    	rs("num")=num
		rs("class")="发布点赠送"
		rs("content")=touser&"得到"&session("useridname")&"会员赠送的"&give1&"个发布点"
		rs("price")=0
		rs("jiegou")="发布点赠送成功"
    	rs.update
    	rs.close
		call check_jieducm_name(session("useridname"))   
    Response.Write("<script>alert('赠送成功!');window.location.href='moneyorpush.asp';</script>")	
   response.End()

elseif action="5" then
	 ReNum1=request.form("ReNum1")
	if isdun=1 then
	 Response.Write("<script>alert('你已被管理员禁止了兑换功能!');history.back();;</script>")
	 response.End()
	elseif int(jifei)<int(ReNum1)*int(stupjifei) then
	 Response.Write("<script>alert('你积分不够,不能兑换!');history.back();</script>")
	 response.End()
	elseif czm<>request("czm") then
	  Response.Write("<script>alert('操作码不正确!');history.back();</script>")
	  response.End()
	 elseif int(ReNum1)<1 then
	 	  Response.Write("<script>alert('输入有误!');history.back();</script>")
	  response.End()
	 	elseif not isNumeric(ReNum1) then
          Response.Write("<script>alert('输入有误请重新输入！');history.back();</script>")
          response.End()
	end if
	
	  Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
	   if rs.eof then
	       Response.Write("<script>alert('无此会员');history.back();</script>")
		   response.End()
	   else
	    rs("jifei")=rs("jifei")-ReNum1*stupjifei
		rs("fabudian")=rs("fabudian")+ReNum1
    	rs.update
    	rs.close
	  end if
	 
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="积分换发布点"
		rs("content")=session("useridname")&"积分换"&ReNum1&"个发布点积分减少了"&ReNum1*stupjifei
		rs("price")=0
		rs("jiegou")="积分换发布点成功"
    	rs.update
    	rs.close	
		call check_jieducm_name(session("useridname"))   	
    Response.Write("<script>alert('兑换成功!');window.location.href='moneyorpush.asp';</script>")
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
function save_onclick()
{
    var Price=document.formt.ToUser.value;
  if(Price=="")
  {
    alert("请输入兑换个数！");
	document.formt.ToUser.focus();
	return false;
	}
if   ((document.formt.ToUser.value.indexOf("-")   ==   0)||!(document.formt.ToUser.value.indexOf(".")   ==   -1)){   
  alert("兑换数量不能为小数或负数");   
  document.formt.ToUser.focus();   
  return   false;   
  } 
if(Price<10)
  {
    alert("10个发布点起兑换！");
	document.formt.ToUser.focus();
	return false;
	}

var code=document.formt.czm.value;
  if(code=="")
  {
    alert("操作码不能为空！");
	document.formt.czm.focus();
	return false;
	}
  return true;  
}


function  save_onclick3()
{

    var ToUser=document.form3.ToUser.value;
  if(ToUser=="")
  {
    alert("请输入要赠送给那位会员！");
	document.form3.ToUser.focus();
	return false;
	}	
	
  var GiveNum=document.form3.GiveNum.value;
  if(GiveNum=="")
  {
    alert("请输入赠送数量！");
	document.form3.GiveNum.focus();
	return false;
	}  


 if(!Number(GiveNum))
	  
  {
    alert("赠送金额必须大于0,请重新输入!");
	document.form3.GiveNum.focus();
	return false;
	}
if   ((document.form3.GiveNum.value.indexOf("-")   ==   0)||!(document.form3.GiveNum.value.indexOf(".")   ==   -1)){   
  alert("赠送金额不能为小数或负数");   
  document.form3.GiveNum.focus();   
  return   false;   
  }  
	

  var code=document.form3.czm.value;	
	if(code=="")
  {
    alert("操作码不能为空！");
	document.form3.czm.focus();
	return false;
	}

  return true;  
}


function  save_onclick5()
{
    var Price=document.form5.ReNum1.value;
  if(Price=="")
  {
    alert("请输入兑换数量！");
	document.form5.ReNum1.focus();
	return false;
	}
 if(!Number(Price))
	  
  {
    alert("请您输入数字!");
	document.form5.ReNum1.focus();
	return false;
	}
if   ((document.form5.ReNum1.value.indexOf("-")   ==   0)||!(document.form5.ReNum1.value.indexOf(".")   ==   -1)){   
  alert("兑换数量不能为小数或负数");   
  document.form5.ReNum1.focus();   
  return   false;   
  } 
 var czm=document.form5.czm.value;
  if(czm=="")
  {
    alert("请输入您的操作码！");
	document.form5.czm.focus();
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
                    <div class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt;我要兑换 &gt;&gt; </div>
                    <div class=pp8><strong>我要兑换</strong></div>
                    <div style="MARGIN-TOP: 10px; LINE-HEIGHT: 200%">
                      <ul>
                        <li>* 如果你推荐的人越多，将获得更多的积分！<br>
* 兑换成功后，系统自动相应的操作。
  <br>
  * 您的操作记录中也会有相应的记录信息 </li>
                      </ul>
                    </div><br><br>
					  <table width="95%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#DEEEFA">
                <tr>
                  <td><STRONG>&nbsp;发布点兑换存款</STRONG></td>
                </tr>
              </table>
              <form action="" method="POST" name="formt" id="formt" onSubmit="return save_onclick()">
                      <table width="639" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="35" align="right"><div align="right">兑换人：</div></td>
                          <td><%=session("useridname")%> 你现在的发布点是：<%=fabudian%> </td>
                        </tr>
                        <tr>
                          <td width="134" height="35" align="right"><div align="right">我想用：</div></td>
                          <td width="505"><input id=ToUser  name=ToUser  onKeyUp="if(isNaN(value))execCommand('undo')">
                            个发布点来兑换(每1个发布点可以兑换
							<%if stupfa<1 then 
							response.Write("0")
							response.Write(stupfa)
                            else
							response.Write(stupfa)
							end if%>元,<%=stupfhcshu%>个发布点起兑换)</td>
                        </tr>
                        <tr>
                          <td height="35"><div align="right">操作码：</div></td>
                          <td><input type="password" name="czm" id="czm"></td>
                        </tr>
                        <tr>
                          <td height="35">&nbsp;</td>
                          <td><input type="submit" name="button" id="button" value="开始兑换" <%if stupfhc=0 then%> disabled="disabled"<%end if%>/>
                              <input type="hidden" name="action" value="1"></td>
                        </tr>
                      </table>
                    </form>

              <table width="95%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#DEEEFA">
                <tr>
                  <td><STRONG>&nbsp;积分换发布点</STRONG></td>
                </tr>
              </table>
              <form name="form5" action="" method="post" onSubmit="return save_onclick5();">
                      <div>
                        <table width="614" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr>
                            <td height="35" align="right"><div align="right">兑换人：</div></td>
                            <td><%=session("useridname")%> 你现在的积分是：<%=jifei%> </td>
                          </tr>
                          <tr>
                            <td width="134" height="35" align="right"><div align="right">我想获得：</div></td>
                            <td width="480"><input id=ReNum1  maxlength=4 name=ReNum1 onKeyUp="if(isNaN(value))execCommand('undo')">
                              个发布点( 每<%=stupjifei%>积分可兑换1个发布点) </td>
                          </tr>
                          <tr>
                            <td height="35"><div align="right">操作码：</div></td>
                            <td><input type="password" name="czm" id="czm"></td>
                          </tr>
                          <tr>
                            <td height="35">&nbsp;</td>
                            <td><input type="submit" name="button" id="button" value="开始兑换" />
                                <input type="hidden" name="action" value="5"></td>
                          </tr>
                        </table>
                      </div>
                    </form>
					
					<table width="95%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#DEEEFA">
                <tr>
                  <td><STRONG>&nbsp;发布点互赠</STRONG></td>
                </tr>
              </table>
              <form name="form3" action="" method="post" onSubmit="return save_onclick3();">
                      <div>
                        <table width="614" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr>
                            <td height="35" align="right"><div align="right">赠送人：</div></td>
                            <td><%=session("useridname")%> 你现在的发布点是：<%=fabudian%> </td>
                          </tr>
                          <tr>
                            <td width="134" height="35" align="right"><div align="right">我要赠送给：</div></td>
                            <td width="480"><input id=ToUser  maxlength=16 name=ToUser >
                              对方帐号 </td>
                          </tr>
                          <tr>
                            <td height="35"><div align="right">赠送数量：</div></td>
                            <td><input id=GiveNum  maxlength=4  name=GiveNum  onKeyUp="if(isNaN(value))execCommand('undo')"></td>
                          </tr>
                          <tr>
                            <td height="35"><div align="right">操作码：</div></td>
                            <td><input type="password" name="czm" id="czm"></td>
                          </tr>
                          <tr>
                            <td height="35">&nbsp;</td>
                            <td><input type="submit" name="button" id="button" value="开始赠送"  <%if isgive=1 then%> disabled="disabled" <%end if%> />
                                <input type="hidden" name="action" value="3"></td>
                          </tr>
                        </table>
                      </div>
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
