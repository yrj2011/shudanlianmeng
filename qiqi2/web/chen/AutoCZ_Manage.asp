<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷����ݶȴ�ý
'�ٷ���ҳ��http://www.jieducm.cn
'����֧�֣�robin 
'QQ;616591415
'�绰��13598006137
'��Ȩ����Ȩ���ڽݶȴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2008��2012 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
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
		 Response.Write("<script>alert('����Ϣ״̬�ѷ����仯!');history.back();</script>")
        response.End()
	 end if 
	 
Sql = "select * from  "&jieducm&"_register where username='"&UserName&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,3,3
if Rs.eof  Then
	Response.Write("<script>alert('�Բ���,������!');history.back();</script>")
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
	Logrs("class") = "֧������ֵ"
	Logrs("content") = UserName & " ���߳�ֵ" & Money & "Ԫ�����׺ţ�" & TradeNo
	Logrs("price") = "+" & Money
	Logrs("jiegou") = "��ֵ�ɹ�"
	Logrs.Update
	Logrs.Close
	Set Logrs = Nothing
	Response.Write("<script>alert('�����ɹ�!');window.location.href='AutoCZ_Manage.asp';</script>")
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
<title>ת�˳�ֵ����</title>
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
* <A href="?">�����Զ���ֵ��¼</A>
* <A href="?action=tianjia">��ӳ�ֵ���׺�</A> 
* <A href="?action=jinzhi">��ӽ�ֹ��ֵ���׺�</A>  
</DIV>
</DIV>
<% If request("action") = "tianjia" Then %>

<div style="margin:0 auto; margin-top:20px; padding:20px; text-align:center; border:1px #09F solid; width:908px; background-color:#F0FCFF">
<FORM name=czform action="?Action=AddNo" method=post><b>��ֵ��Ա�˺ţ�</b><input name=UserName id=UserName type=text class=textbox size=20 maxlength=20 /><br><br>
<b> &nbsp; ֧�����׺ţ�</b><input name=TradeNo id=TradeNo type=text class=textbox size=20 maxlength=38 /><br><br>
<b> &nbsp; &nbsp; ��ֵ��</b><input name=Money id=Money type=text class=textbox size=20 maxlength=8 /><br><div style="text-align:center;"><br><Input class=button id=Submit type=Submit value=ȷ���ύ name=Submit /></div></form>
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
<b>��ӽ�ֹ��ֵ���׺���ҪĿ�ı��������������Ӷ����ֵ��</b><br><br>1������Ա������ĳ֧�����ţ�ƽ̨֧������ֵҲ�Ǹ�֧�����ŵĻ�����Ա������Ա��깺������Ʒ�󣬰�֧�������׵��ͽ����д��ƽ̨���г�ֵ����õ���Ӧ�ʽ�<br><br>2��ĳ�ͻ�ͨ������;�����㽻�׻����տ����֧����֧����һ���ʽ𣬹�����ƽ̨��֧������ֵ���ܣ��������뽻�׺źͽ���ȥ��ֵ��<br><br>���Խ���ʹ������Զ���ֵ�����ѣ�ע����µ�֧�����˺��ٹ�����֮ǰ�Ѿ���֤��֧������ʹ����֧�����˺š�<br><br><FORM name=czform action="?Action=SetNo" method=post><b>�������ֹ��ֵ�Ľ��׺ţ�</b><input name=TradeNo id=TradeNo type=text class=textbox size=20 maxlength=38 /> &nbsp; <Input class=button id=Submit type=Submit value=ȷ���ύ name=Submit /></form></div>

<% Else %>
<table border=0 width=100% align=center bgcolor="#FFFFFF">
  <tr>
            <td width="100%"> 
              <div align="center">
                <form name="form1" method="post" action="?action=jieducm_ok">
                  
                  �������׺ţ�
                  <input class=input1 size=25 name=text value="<%=request("text")%>">
��
<input name="submit" type=submit class=input2 value=" �� �� ">
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
  <LI style="WIDTH: 10%">��Ա��</LI>
  <LI style="WIDTH: 18%">���׺�</LI>
  <LI style="WIDTH: 10%">��ֵ���</LI>
  <LI style="WIDTH: 12%">�ύʱ��</LI>
  <LI style="WIDTH: 12%">���ʱ��</LI>
  <LI style="WIDTH: 12%">IP</LI>
  <LI style="WIDTH: 10%">״̬</LI>
   <LI style="WIDTH: 8%">����</LI>
  
  </UL></DIV></DIV>
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid">
<form name="myform" method="Post" action="?" onSubmit="return(confirm('ȷ��Ҫɾ����'))">
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
		  response.Write("�Բ���û������Ҫ��ҳ��")
          Response.End
	    end if
	 end if
end if    
if PageNo<0 then
	response.Write("û����һҳ!")
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
		Response.Write "<span style=""color:green"">�ȴ���ֵ</span>"
	ElseIf rs("Status")=1 Then
		Response.Write "��ֵ�ɹ�"
	ElseIf rs("Status")=2 Then
		Response.Write "<span style=""color:red"">��ֵʧ��</span>"
	ElseIf rs("Status")=3 Then
		Response.Write "<span style=""color:blue"">��ֹ��ֵ</span>"
	End If
	%></LI>
	<LI style="WIDTH: 8%">
	<%If rs("Status")=0 Then%>
	<a href="?action=ok&id=<%=rs("id")%>" onClick="return confirm('ȷ����ֵ����Զ�ת����Ա�˺�����');"><span style="color:red">ȷ�ϳ�ֵ</span></a>
	<%End If%>
	
	</LI>
	
	</UL></DIV>
   <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
<DIV style="WIDTH: 98%; LINE-HEIGHT: 40px; PADDING-TOP: 10px; HEIGHT: 40px; TEXT-ALIGN: center"> <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"����¼") %></DIV>
<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" style="border: 0px;background-color: #eeeeee;">
              ѡ�б�ҳ��ʾ�����м�¼
<input name="submit" type='submit' value='&nbsp;ɾ��ѡ���ļ�¼&nbsp;' style="cursor: hand;background-color: #cccccc;">
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
		Call AlertHistory("�����뽻�׺ţ�",-1)
		Exit Sub
	end if

	Dim RS: Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open "SELECT * FROM AutoCZ WHERE TradeNo = '" & TradeNo & "'", Conn, 1, 3
	If Not Rs.EOF Then
		Rs.Close: Set Rs = Nothing
		Call AlertHistory("�ý��׺��Ѿ����ڣ�",-1)		
	Else
		Rs.AddNew
		Rs("UserName") =  AdminName
		RS("RealName") = "����Ա��̨����"
		RS("TradeNo") = TradeNo
		Rs("Money") = 0
		Rs("Status") = 3
		Rs("AddDate") = Now
		Rs("EndDate") = Now
		Rs("IP") = Request.ServerVariables("REMOTE_ADDR")
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Response.Write("<script language=""Javascript"">alert('��ӳɹ���');window.location.href='AutoCZ_Manage.asp';</script>")	
	End If
End Sub

Sub DoSaveAlipPayCZ()
	Dim TradeNo: TradeNo=Left(Trim(Request("TradeNo")),28)
	Dim Money:Money=Left(Trim(Request("Money")),8)
	jieducm_Len=Len(TradeNo)
	If TradeNo="" Or Money="" Then 
		Call AlertHistory("������֧�����Ľ��׺Ż�֧����",-1)
		Exit Sub
	end if
 
	If Not IsNumeric(Money) Then
		Call AlertHistory("֧���Ľ����������",-1)
		Exit Sub
	End if
	If Round(Money,2) <= 0 Then
		Call AlertHistory("֧���Ľ������Ǵ���0",-1)
		Exit Sub
	End if
	IF Conn.Execute("SELECT ID FROM " & jieducm & "_register WHERE UserName='" & Trim(Request("UserName")) & "'").Eof Then
		Call AlertHistory("������ĳ�ֵ��Ա�˺Ų����ڣ�",-1)
		Exit Sub
	End IF
	Dim RS: Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open "SELECT * FROM AutoCZ WHERE TradeNo = '" & TradeNo & "'", Conn, 1, 3
	If Not Rs.EOF Then
		If Rs.RecordCount >= 2 then
			Rs.Close: Set Rs = Nothing
			Call AlertHistory("��Ǹ��ϵͳ����ͬһ���׺Ų����ύ����2��������ɾ����������ӣ�лл��",-1)			
		ElseIf Rs("Status") = 0 Then
			Rs.Close: Set Rs = Nothing
			Call AlertHistory("�ý��׺��Ѿ����ڣ��벻Ҫ�ظ��ύ��лл��",-1)			
		ElseIf Rs("Status") = 1 Then
			Rs.Close: Set Rs = Nothing
			Call AlertHistory("�ý��׺��Ѿ���ֵ�ɹ����벻Ҫ�ظ��ύ��лл��",-1)			
		ElseIf Rs("Status") = 3 Then
			Rs.Close: Set Rs = Nothing
			Call AlertHistory("�ý��׺�ϵͳ�ѽ�ֹ��ֵ������ɾ����������ӣ�лл��",-1)			
		End IF
	End If
	'��������ͬ���׺ŵ� �� ��һ�γ�ֵʧ��(Status=2)���������ύһ��
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
	Response.Write("<script language=""Javascript"">alert('�ύ�ɹ���\n\nϵͳ������1�������ҳ�ֵ��ϣ�');window.location.href='AutoCZ_Manage.asp';</script>")	
End Sub

Sub DoDel()
	if Request("ID") <> "" then
		Conn.Execute ("DELETE FROM AutoCZ Where ID IN(" & Request("ID") & ")")
		Response.Write("<script language=""Javascript"">alert('ɾ���ɹ�');window.location.href='" & Request.ServerVariables("HTTP_REFERER") & "';</script>")
	end if
End Sub
%>