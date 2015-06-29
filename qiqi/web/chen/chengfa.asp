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
name1=request("username")
if request.Form<>"" then

 username=request("name")
 sf=request("sf")
 class1=request("class")
 fabudian=Replace(Trim(request.Form("fabudian")),"'","''")
 content=request("content")
 czm=request("czm")
 if sf=1 then
  leibie="官方奖励"
 else
  leibie="官方处罚"
 end if
 
if md5(czm)<>admin_czmpass then
   Set rs=server.createobject("ADODB.RECORDSET")
   rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
   rs.addnew
   rs("username")=session("AdminName")
   rs("class")=leibie
   rs("content")=session("AdminName")&"管理员给"&username&"进行"&leibie&"时操作码输入错误"
   rs("jiegou")="失败"
   rs.update
   rs.close
   Response.Write("<script>alert('操作码有误!请不要重复操作，平台会记录你的行为！');history.back();</script>")
   response.End()
end if

if sf=2 then
  if class1="1" then
   sqlr="update "&jieducm&"_register set cunkuan=cunkuan-"&fabudian&" where username='"&username&"'"
  elseif class1="2"  then
     sqlr="update "&jieducm&"_register set fabudian=fabudian-"&fabudian&" where username='"&username&"'"
  elseif class1="3"  then
       sqlr="update "&jieducm&"_register set jifei=jifei-"&fabudian&" where username='"&username&"'"
  end if
elseif sf=1 then
  if class1="1" then
   sqlr="update "&jieducm&"_register set cunkuan=cunkuan+"&fabudian&" where username='"&username&"'"
  elseif class1="2"  then
     sqlr="update "&jieducm&"_register set fabudian=fabudian+"&fabudian&" where username='"&username&"'"
  elseif class1="3"  then
       sqlr="update "&jieducm&"_register set jifei=jifei+"&fabudian&" where username='"&username&"'"
  end if
end if
conn.execute sqlr
  

if class1="1" then
classname="元存款"
elseif class1="2" then
classname="个发布点"
elseif class1="3" then
classname="个积分" 
end if  

 	 Set rs=server.createobject("ADODB.RECORDSET")
     rs.open "Select * From "&jieducm&"_chengfa" ,Conn,3,3  
      rs.addnew
	  rs("username")=username
	  rs("class")=class1
	  rs("num")=fabudian
	  rs("now")=now()
	  rs("manage")=session("AdminName")
	  rs("content")=content
	  rs("leibie")=sf
	  rs.update
	  rs.close
	  
 	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=username
    	rs("num")=num
		rs("class")=leibie
		rs("content")=session("AdminName")&"管理员"&leibie&""&username&"了"&fabudian&""&classname
		if class1=1 then
        if sf=1 then
         rs("price")="+"&fabudian
        elseif sf=2 then
		rs("price")="-"&fabudian
         end if
		else
		rs("price")=0
		end if
		rs("jiegou")="操作成功"
    	rs.update
    	rs.close
		
	Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
		rs("class")=leibie
		rs("content")=session("AdminName")&"管理员"&leibie&""&username&"了"&fabudian&""&classname
		rs("jiegou")=leibie&"惩罚成功"
    	rs.update
    	rs.close		
   Response.Write("<script>alert('操作成功!');window.location.href='usermanage.asp?action=sear&text="&username&"';</script>")
   response.End()
end if
%>
<LINK href="../images/Default.css" type=text/css rel=stylesheet>
<LINK href="../images/Bbs.css" type=text/css rel=stylesheet>
<LINK href="../images/style.css" type=text/css rel=stylesheet>
<script language=javascript>
<!--
function CheckForm()
{
	if(document.form.fabudian.value=="")
	{
		alert("请输入金额！");
		document.form.fabudian.focus();
		return false;
	}
  if(document.form.czm.value=="")
	{
		alert("请输入操作码！");
		document.form.czm.focus();
		return false;
	}

}
//-->
</script>
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 10px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid">
<DIV> </DIV>

<DIV style="WIDTH: 95%; PADDING-TOP: 10px; HEIGHT: 160px">
<DIV style="PADDING-LEFT: 50px; WIDTH: 98%; HEIGHT: 20px">
  <div align="center"><SPAN 
style="FONT-WEIGHT: bolder; FONT-SIZE: 14px; COLOR: red">官方赏罚管理</SPAN><BR>
  </div>
</DIV>
<UL class=pushmission>
  <LI><FORM id=form name=form  action="" method="post" onSubmit="return CheckForm();">

    <table width="100%" border="0">
      <tr>
        <td width="47%"><div align="right">惩罚用户名：</div></td>
        <td width="53%"><input type="text" name="name" id="name"  value="<%=name1%>" readonly></td>
      </tr>
      <tr>
        <td><div align="right">类别：</div></td>
        <td><label>
          <input name="sf" type="radio" value="1" checked="checked" />
          奖励
          <input type="radio" name="sf" value="2" />处罚
        </label></td>
      </tr>
      <tr>
        <td><div align="right">操作类别:</div></td>
        <td><select name="class" id="class">
          <option value="1" selected>存款</option>
          <option value="2">发布点</option>
          <option value="3">积分</option>
          </select>          </td>
      </tr>
      <tr>
        <td><div align="right">金额：</div></td>
        <td><input name="fabudian" type="text" id="fabudian" onkeyup="if(isNaN(value))execCommand('undo')" maxlength="4"></td>
      </tr>
	       <tr>
        <td><div align="right">原因：</div></td>
        <td><label>
          <textarea name="content" rows="5"></textarea>
          </label></td>
      </tr>
	    <tr>
        <td><div align="right">操作码：</div></td>
        <td><label>
        <input type="password" name="czm" />
        </label></td>
      </tr>
      <tr>
        <td colspan="2"><div align="center">
          <input type="submit" name="button" id="button" value="提 交">
          &nbsp;&nbsp;&nbsp;&nbsp; 
          <input type="button" name="button2" id="button2" value="取 消"  onClick="history.back();"/>
        </div></td>
        </tr>
    </table></FORM>
  </LI>
</UL>
</DIV>

</DIV>
<%
set rs=nothing
conn.close
set conn=nothing%>