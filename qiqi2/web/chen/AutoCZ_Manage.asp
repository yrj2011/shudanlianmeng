<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
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
'Copyright (C) 2008－2012 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
action=request.QueryString("action")
id=request.QueryString("id")
if action="ok" then

Sql= "select * from AutoCZ where id="&id&" and Status=0"
	 Set rs=server.createobject("ADODB.RECORDSET")
	 Rs.Open Sql,conn,3,3
	 if not rs.eof then
	    UserName=rs("UserName")
		Money=rs("Money")
		TradeNo=rs("TradeNo")
		rs("Status")=1
		rs("EndDate")=now()
		rs.update
		rs.close
	else
		 Response.Write("<script>alert('此信息状态已发生变化!');history.back();</script>")
        response.End()
	 end if 
	 
Sql = "select * from  "&jieducm&"_register where username='"&UserName&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,3,3
if Rs.eof  Then
	Response.Write("<script>alert('对不起,出错了!');history.back();</script>")
	response.End()
Else
rs("cunkuan")=rs("cunkuan")+Money
rs.update
rs.close
end if


Randomize
	Dim num: num=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&int(rnd*100)
	Dim Logrs: Set Logrs=server.createobject("ADODB.RECORDSET")
	Logrs.Open "Select * From " &jieducm& "_record " , Conn, 1, 3 
	Logrs.AddNew
	Logrs("username") = UserName
	Logrs("num") = num
	Logrs("class") = "支付宝充值"
	Logrs("content") = UserName & " 在线充值" & Money & "元，交易号：" & TradeNo
	Logrs("price") = "+" & Money
	Logrs("jiegou") = "充值成功"
	Logrs.Update
	Logrs.Close
	Set Logrs = Nothing
	Response.Write("<script>alert('操作成功!');window.location.href='AutoCZ_Manage.asp';</script>")
	response.End()	
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<LINK href="../images/mission.css" type=text/css rel=stylesheet>
<LINK href="../images/Bbs.css" type=text/css rel=stylesheet>
<LINK href="../images/style.css" type=text/css rel=stylesheet>
<title>转账充值管理</title>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll"&&e.disabled==false)
       e.checked = form.chkAll.checked;
    }
  }

</SCRIPT>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   

<DIV 
style="CLEAR: both; OVERFLOW: hidden; WIDTH: 950px; BACKGROUND-COLOR: #f3f8fe">
  <DIV style="CLEAR: both; MARGIN-BOTTOM: 10px; MARGIN-LEFT: 1px;MARGIN-TOP: 10px; COLOR: #ff0000; text-align:center; font-weight:bold">
* <A href="?">所有自动充值记录</A>
* <A href="?action=tianjia">添加充值交易号</A> 
* <A href="?action=jinzhi">添加禁止充值交易号</A>  
</DIV>
</DIV>
<% If request("action") = "tianjia" Then %>

<div style="margin:0 auto; margin-top:20px; padding:20px; text-align:center; border:1px #09F solid; width:908px; background-color:#F0FCFF">
<FORM name=czform action="?Action=AddNo" method=post><b>充值会员账号：</b><input name=UserName id=UserName type=text class=textbox size=20 maxlength=20 /><br><br>
<b> &nbsp; 支付交易号：</b><input name=TradeNo id=TradeNo type=text class=textbox size=20 maxlength=38 /><br><br>
<b> &nbsp; &nbsp; 充值金额：</b><input name=Money id=Money type=text class=textbox size=20 maxlength=8 /><br><div style="text-align:center;"><br><Input class=button id=Submit type=Submit value=确定提交 name=Submit /></div></form>
</div>
<%
ElseIF request("action") = "AddNo" Then
	Call DoSaveAlipPayCZ()
ElseIF request("action") = "SetNo" Then
	Call DoSetNo()
ElseIF request("action") = "Del" Then
	Call DoDel()
ElseIF request("action") = "jinzhi" Then %>
<div style="margin:0 auto; margin-top:20px; padding:20px; text-align:center; border:1px #09F solid; width:908px; background-color:#F0FCFF">
<b>添加禁止充值交易号主要目的避免如下类似例子恶意充值：</b><br><br>1、你的淘宝店绑定了某支付宝号，平台支付宝充值也是该支付宝号的话，会员在你的淘宝店购买了商品后，把支付宝交易单和金额填写到平台进行充值，会得到相应资金；<br><br>2、某客户通过其他途径与你交易或是收款，向你支付宝支付了一笔资金，过后发现平台有支付宝充值功能，于是输入交易号和金额进去充值。<br><br>所以建议使用软件自动充值的朋友，注册个新的支付宝账号再关联到之前已经认证的支付宝，使用新支付宝账号。<br><br><FORM name=czform action="?Action=SetNo" method=post><b>请输入禁止充值的交易号：</b><input name=TradeNo id=TradeNo type=text class=textbox size=20 maxlength=38 /> &nbsp; <Input class=button id=Submit type=Submit value=确定提交 name=Submit /></form></div>

<% Else %>
<table border=0 width=100% align=center bgcolor="#FFFFFF">
  <tr>
            <td width="100%"> 
              <div align="center">
                <form name="form1" method="post" action="?action=jieducm_ok">
                  
                  搜索交易号：
                  <input class=input1 size=25 name=text value="<%=request("text")%>">
　
<input name="submit" type=submit class=input2 value=" 搜 索 ">
                </form> 
            </div></td>
  </tr>
          <tr>
            <td height=12>&nbsp;</td>
          </tr>
          </td>
</table>
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #e6f7fb">
<DIV style="CLEAR: both; WIDTH: 98%; LINE-HEIGHT: 35px; HEIGHT: 35px">
<UL class=mission>
  <LI style="WIDTH: 8%">ID</LI>
  <LI style="WIDTH: 10%">会员名</LI>
  <LI style="WIDTH: 18%">交易号</LI>
  <LI style="WIDTH: 10%">充值金额</LI>
  <LI style="WIDTH: 12%">提交时间</LI>
  <LI style="WIDTH: 12%">完成时间</LI>
  <LI style="WIDTH: 12%">IP</LI>
  <LI style="WIDTH: 10%">状态</LI>
   <LI style="WIDTH: 8%">操作</LI>
  
  </UL></DIV></DIV>
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid">
<form name="myform" method="Post" action="?" onSubmit="return(confirm('确定要删除吗？'))">
<%	
action=request("action")

if action="" then
	sql="select * from AutoCZ  order by id desc"
elseif action="jieducm_ok" then
	sql="select * from AutoCZ where TradeNo like '%"& request("text") &"%' order by id desc"
elseif action="search" then
	sql="select * from AutoCZ where UserName like '%"& request("UserName") &"%' order by id desc"
end if
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
	dim maxperpage,url,PageNo
	if action="" then
	   url="?"
	else
		url="?action="& action
  	end if
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
<DIV style="WIDTH: 98%; PADDING-TOP: 10px; BORDER-BOTTOM: #06314a 1px dashed; HEIGHT: 30px">
<UL class=missionitem>
 
  <LI style="WIDTH: 8%"><input name='ID' type='checkbox' onClick="unselectall()" id="ID" style="border: 0px;background-color: #eeeeee;" value=<%=rs("id")%>><%=rs("id")%> </LI>
  <LI style="WIDTH: 10%"><a href="?action=search&UserName=<%=rs("UserName")%>"><%=rs("UserName")%></a></LI>
  <LI style="WIDTH: 18%"><%=rs("TradeNo")%></LI>
  <LI style="WIDTH: 10%; font-weight:bold"><%=rs("Money")%></LI>
  <LI style="WIDTH: 12%"><%=rs("AddDate")%></LI>
  <LI style="WIDTH: 12%"><%if rs("EndDate") <>"" then response.write rs("EndDate") else response.write "&nbsp;"%></LI>
  <LI style="WIDTH: 12%"><%=rs("IP")%></LI>
  <LI style="WIDTH: 10%"><%
	If rs("Status")=0 Then
		Response.Write "<span style=""color:green"">等待充值</span>"
	ElseIf rs("Status")=1 Then
		Response.Write "充值成功"
	ElseIf rs("Status")=2 Then
		Response.Write "<span style=""color:red"">充值失败</span>"
	ElseIf rs("Status")=3 Then
		Response.Write "<span style=""color:blue"">禁止充值</span>"
	End If
	%></LI>
	<LI style="WIDTH: 8%">
	<%If rs("Status")=0 Then%>
	<a href="?action=ok&id=<%=rs("id")%>" onClick="return confirm('确定充值后存款将自动转到会员账号里面');"><span style="color:red">确认充值</span></a>
	<%End If%>
	
	</LI>
	
	</UL></DIV>
   <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
<DIV style="WIDTH: 98%; LINE-HEIGHT: 40px; PADDING-TOP: 10px; HEIGHT: 40px; TEXT-ALIGN: center"> <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条记录") %></DIV>
<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" style="border: 0px;background-color: #eeeeee;">
              选中本页显示的所有记录
<input name="submit" type='submit' value='&nbsp;删除选定的记录&nbsp;' style="cursor: hand;background-color: #cccccc;">
              <input name="Action" type="hidden" id="Action" value="Del"></form>
&nbsp;&nbsp;&nbsp;

</DIV>

</BODY></HTML>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
End IF

Function AlertHistory(SuccessStr, N)
	Response.Write("<script language=""Javascript""> alert('" & SuccessStr & "');history.back(" & N & ");</script>")
	Response.End()
End Function

Sub DoSetNo()
	Dim TradeNo: TradeNo=Left(Trim(Request("TradeNo")),38)
	If TradeNo="" Then 
		Call AlertHistory("请输入交易号！",-1)
		Exit Sub
	end if

	Dim RS: Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open "SELECT * FROM AutoCZ WHERE TradeNo = '" & TradeNo & "'", Conn, 1, 3
	If Not Rs.EOF Then
		Rs.Close: Set Rs = Nothing
		Call AlertHistory("该交易号已经存在！",-1)		
	Else
		Rs.AddNew
		Rs("UserName") =  AdminName
		RS("RealName") = "管理员后台操作"
		RS("TradeNo") = TradeNo
		Rs("Money") = 0
		Rs("Status") = 3
		Rs("AddDate") = Now
		Rs("EndDate") = Now
		Rs("IP") = Request.ServerVariables("REMOTE_ADDR")
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Response.Write("<script language=""Javascript"">alert('添加成功！');window.location.href='AutoCZ_Manage.asp';</script>")	
	End If
End Sub

Sub DoSaveAlipPayCZ()
	Dim TradeNo: TradeNo=Left(Trim(Request("TradeNo")),28)
	Dim Money:Money=Left(Trim(Request("Money")),8)
	jieducm_Len=Len(TradeNo)
	If TradeNo="" Or Money="" Then 
		Call AlertHistory("请输入支付宝的交易号或支付金额！",-1)
		Exit Sub
	end if
 
	If Not IsNumeric(Money) Then
		Call AlertHistory("支付的金额输入有误！",-1)
		Exit Sub
	End if
	If Round(Money,2) <= 0 Then
		Call AlertHistory("支付的金额必需是大于0",-1)
		Exit Sub
	End if
	IF Conn.Execute("SELECT ID FROM " & jieducm & "_register WHERE UserName='" & Trim(Request("UserName")) & "'").Eof Then
		Call AlertHistory("你输入的充值会员账号不存在！",-1)
		Exit Sub
	End IF
	Dim RS: Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open "SELECT * FROM AutoCZ WHERE TradeNo = '" & TradeNo & "'", Conn, 1, 3
	If Not Rs.EOF Then
		If Rs.RecordCount >= 2 then
			Rs.Close: Set Rs = Nothing
			Call AlertHistory("抱歉！系统限制同一交易号不能提交超过2条！请先删除再重新添加！谢谢！",-1)			
		ElseIf Rs("Status") = 0 Then
			Rs.Close: Set Rs = Nothing
			Call AlertHistory("该交易号已经存在，请不要重复提交！谢谢！",-1)			
		ElseIf Rs("Status") = 1 Then
			Rs.Close: Set Rs = Nothing
			Call AlertHistory("该交易号已经充值成功，请不要重复提交！谢谢！",-1)			
		ElseIf Rs("Status") = 3 Then
			Rs.Close: Set Rs = Nothing
			Call AlertHistory("该交易号系统已禁止充值，请先删除再重新添加！谢谢！",-1)			
		End IF
	End If
	'不存在相同交易号的 和 第一次充值失败(Status=2)的允许再提交一次
	Rs.AddNew
	Rs("UserName") = Trim(Request("UserName"))
	RS("RealName") = ""
	RS("TradeNo") = TradeNo
	Rs("Money") = Round(Money,2)
	Rs("Status") = 0
	Rs("AddDate") = Now
	Rs("IP") = Request.ServerVariables("REMOTE_ADDR")
	Rs.Update
	Rs.Close
	Set Rs = Nothing
	Response.Write("<script language=""Javascript"">alert('提交成功！\n\n系统将会在1分钟左右充值完毕！');window.location.href='AutoCZ_Manage.asp';</script>")	
End Sub

Sub DoDel()
	if Request("ID") <> "" then
		Conn.Execute ("DELETE FROM AutoCZ Where ID IN(" & Request("ID") & ")")
		Response.Write("<script language=""Javascript"">alert('删除成功');window.location.href='" & Request.ServerVariables("HTTP_REFERER") & "';</script>")
	end if
End Sub
%>