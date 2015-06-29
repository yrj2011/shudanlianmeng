 <!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="../md5.asp"-->
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
'Copyright (C) 2008－2011 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
 id=request("id")
if request.Form<>"" then
 username=request("username")
 cunkuan=request("cunkuan")
 fabudian=request("fabudian")
 jifei=request("jifei")
 class1=request("class")
 content=request("content")
 czm=request("czm")

IF not isNumeric(request("cunkuan")) then
   Response.Write("<script>alert('操作存款必须为数字，请重新输入！');history.back();</script>")
   response.End()
elseif not isNumeric(request("fabudian")) then
   Response.Write("<script>alert('操作发布点必须为数字，请重新输入！');history.back();</script>")
   response.End()
elseif not isNumeric(request("jifei")) then
   Response.Write("<script>alert('操作积分必须为数字，请重新输入！');history.back();</script>")
   response.End()
end if


if md5(czm)<>admin_czmpass then
   Set rs=server.createobject("ADODB.RECORDSET")
   rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
   rs.addnew
   rs("username")=session("AdminName")
   rs("class")=class1
   rs("content")=session("AdminName")&"管理员给"&username&"进行后台充值操作码输入错误"
   rs("jiegou")="失败"
   rs.update
   rs.close
   
     num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=username
    	rs("num")=num
		rs("class")=class1
		rs("content")=session("AdminName")&"管理员给"&username&"进行后台充值操作码输入错误"
        rs("price")=0	
		rs("jiegou")="失败"
		rs("now")=date()
    	rs.update
    	rs.close
   Response.Write("<script>alert('操作码有误!请不要重复操作，平台会记录你的行为！');history.back();</script>")
   response.End()
end if

if cunkuan="" then
 cunkuan=0
end if
if fabudian="" then
 fabudian=0
end if
if jifei="" then
 jifei=0
end if 



 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select cunkuan,fabudian,jifei From "&jieducm&"_register where username='"&username&"'" ,Conn,3,3  
		rs("cunkuan")=rs("cunkuan")+cunkuan
		rs("fabudian")=rs("fabudian")+fabudian
		rs("jifei")=rs("jifei")+jifei
    	rs.update
    	rs.close
  
if cunkuan<>0 then
  dian=cunkuan&"元存款"
end if
if fabudian<>0 then
  dian=dian&fabudian&"个发布点"
end if
if jifei<>0 then
  dian=dian&jifei&"个积分"
end if

  	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=username
    	rs("num")=num
		rs("class")=class1
		rs("content")=session("AdminName")&"管理员给"&username&"充了"&dian
        rs("price")=cunkuan
		rs("jiegou")="成功"
		rs("now")=date()
    	rs.update
    	rs.close
	
	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_chongjilu" ,Conn,3,3  
	    rs.addnew
		rs("class")="后台充值"
		rs("username")=username
    	rs("fabudian")=fabudian
		rs("cunkuan")=cunkuan
		rs("jifei")=jifei
		rs("now")=now()
    	rs.update
    	rs.close  
		
	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
		rs("class")=class1
		rs("content")=session("AdminName")&"管理员给"&username&"充了"&dian
		rs("jiegou")="成功"		
    	rs.update
    	rs.close
   Response.Write("<script>alert('充值成功!');window.location.href='usermanage.asp?action=sear&text="&username&"';</script>")
end if
 Set rs=server.createobject("ADODB.RECORDSET")
 rs.open "Select * From "&jieducm&"_register where id="&id&"",Conn,3,3
 if not rs.eof then
%>
<LINK href="../images/Default.css" type=text/css rel=stylesheet>
<LINK href="../images/Bbs.css" type=text/css rel=stylesheet>
<LINK href="../images/style.css" type=text/css rel=stylesheet>
<SCRIPT language=javascript>
function  save_onclick22()
{

if   (!(document.aspnetForm.fabudian.value.indexOf(".")   ==   -1)){   
  alert("操作发布点不能为小数");   
  document.aspnetForm.fabudian.focus();   
  return   false;   
  }	
if   (!(document.aspnetForm.cang.value.indexOf(".")   ==   -1)){   
  alert("操作收藏点不能为小数");   
  document.aspnetForm.cang.focus();   
  return   false;   
  }	
  if   (!(document.aspnetForm.jifei.value.indexOf(".")   ==   -1)){   
  alert("操作积分不能为小数");   
  document.aspnetForm.jifei.focus();   
  return   false;   
  }	
  return true;  
}
 
  function   test(str){   
          var   pos;   
          var   fst   
          var   lst;   
          if   (str   ==   "")   return;   
          pos   =   str.indexOf(".");   
          if   (pos   !=   -1){   
                  fst   =   str.substring(0,pos);   
                  lst   =   str.substring(pos+1,pos.length);   
                  if   (lst.length   >   1){                             
                            document.aspnetForm.cunkuan.value="";   
							  alert("操作存款不能有两位小数");   

                  }   
          }           
  }  
  function   test1(str){   
          var   pos;   
          var   fst   
          var   lst;   
          if   (str   ==   "")   return;   
          pos   =   str.indexOf(".");   
          if   (pos   !=   -1){   
                  fst   =   str.substring(0,pos);   
                  lst   =   str.substring(pos+1,pos.length);   
                  if   (lst.length   >   1){                             
                            document.aspnetForm.fabudian.value="";   
							  alert("操作发布点不能有两位小数");   

                  }   
          }           
  }  
  function   test2(str){   
          var   pos;   
          var   fst   
          var   lst;   
          if   (str   ==   "")   return;   
          pos   =   str.indexOf(".");   
          if   (pos   !=   -1){   
                  fst   =   str.substring(0,pos);   
                  lst   =   str.substring(pos+1,pos.length);   
                  if   (lst.length   >   1){                             
                            document.aspnetForm.cang.value="";   
							  alert("操作收藏点不能有两位小数");   

                  }   
          }           
  }  
  function   test3(str){   
          var   pos;   
          var   fst   
          var   lst;   
          if   (str   ==   "")   return;   
          pos   =   str.indexOf(".");   
          if   (pos   !=   -1){   
                  fst   =   str.substring(0,pos);   
                  lst   =   str.substring(pos+1,pos.length);   
                  if   (lst.length   >   1){                             
                            document.aspnetForm.jifei.value="";   
							  alert("操作积分不能有两位小数");   

                  }   
          }           
  }   
  </script>   
    
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 10px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid">
<FORM id=aspnetForm name=aspnetForm  action="" method=post  onSubmit="return save_onclick22();">
<DIV> </DIV>
<DIV style="WIDTH: 98%; PADDING-TOP: 10px"><SPAN 
id=ctl00_AllBaseContentPlaceHolder_MsgLabel style="COLOR: red"></SPAN><SPAN 
id=ctl00_AllBaseContentPlaceHolder_RequiredFieldValidator1 
style="DISPLAY: none; COLOR: red">* 充值金额不能为空！</SPAN> <SPAN 
id=ctl00_AllBaseContentPlaceHolder_RangeValidator1 
style="DISPLAY: none; COLOR: red">* 充值金额范围在0.01-999之间！</SPAN> <SPAN 
id=ctl00_AllBaseContentPlaceHolder_RequiredFieldValidator2 
style="DISPLAY: none; COLOR: red">* 充值用户名不能为空！</SPAN></DIV>
<DIV style="WIDTH: 95%; PADDING-TOP: 10px; HEIGHT: 160px">
<DIV style="PADDING-LEFT: 50px; WIDTH: 98%; HEIGHT: 20px">
  <div align="center"><SPAN 
style="FONT-WEIGHT: bolder; FONT-SIZE: 14px; COLOR: red">在线充值：立即到账</SPAN><BR>
  </div>
</DIV>
<UL class=pushmission>
  <LI>
    <table width="100%" border="0">
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td width="39%"><div align="right">充值用户名：</div></td>
        <td width="61%"><input type="text" name="name" id="name"  value="<%=rs("username")%>" readonly>
          存款：<%
if rs("cunkuan")=0 then
response.Write("0.0")
else
response.Write(formatnumber(rs("cunkuan"),1))
end if
%> 发布点：<%
if rs("fabudian")=0 then
response.Write("0.0")
else
response.Write(formatnumber(rs("fabudian"),1))
end if
%>            积分：<%=rs("jifei")%></td>
      </tr>
      <tr>
        <td><div align="right">操作存款：</div></td>
        <td>
        <INPUT id=cunkuan    name=cunkuan maxLength=4    value="0"  onkeyup="test(this.value)"   />
              
          <span class="STYLE1">X或-X 必须是数字</span></td>
      </tr>
      <tr>
        <td><div align="right">操作发布点：</div></td>
        <td><input name="fabudian" type="text" id="fabudian"   maxLength=4   value="0"  onkeyup="test1(this.value)"/>
          <span class="STYLE1">X或-X 必须是数字</span></td>
      </tr>
      
    
      <tr>
        <td><div align="right">操作积分：</div></td>
        <td><input name="jifei" type="text" id="jifei"   maxLength=5  value="0"  onkeyup="test3(this.value)"/>
          <span class="STYLE1">X或-X 必须是数字</span></td>
      </tr> 
      
       <tr>
        <td><div align="right">操作类型：</div></td>
        <td><input name="class" type="radio" id="class" value="任务" checked="checked" />
          任务 <span class="STYLE1">*比便刷任务时出差错需要修改</span></td>
      </tr> 
       <tr>
         <td>&nbsp;</td>
         <td><input type="radio" name="class" id="class" value="充值" />
           充值 <span class="STYLE1">*用户真实充值比如在线充值失败，支付宝充值。</span></td>
       </tr>
       <tr>
         <td>&nbsp;</td>
         <td><input type="radio" name="class" id="class" value="增值" />
           增值 <span class="STYLE1">*用户真实增值比如购买发布点，兑换发布点，换存款。</span></td>
       </tr>
       <tr>
         <td>&nbsp;</td>
         <td><input type="radio" name="class" id="class" value="提现" />
           提现 <span class="STYLE1">*几乎不用。</span></td>
       </tr>
       <tr>
         <td>&nbsp;</td>
         <td><input type="radio" name="class" id="class" value="其它" />
           其它 <span class="STYLE1">*比如自己人充值，送发布点积分给用户。</span></td>
       </tr>
       
        <tr>
         <td><div align="right">操作码：</div></td>
         <td><label>
           <input type="password" name="czm" />
         </label></td>
       </tr>
       <tr>
         <td><div align="right">理由:</div></td>
         <td><textarea name="content" id="content" cols="45" rows="5"></textarea></td>
       </tr>
      <tr>
        <td colspan="2"><div align="center">
         <input type="hidden" name="username" value="<%=rs("username")%>" />
		  <input type="submit" name="button" id="button" value="提 交">
          &nbsp; 
          <input type="button" name="button2" id="button2" value="取 消"  onClick="history.back();"/>
        </div></td>
        </tr>
    </table>
  </LI>
</UL>
</DIV>


<DIV></DIV>

</FORM></DIV>
<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
<% end if%>