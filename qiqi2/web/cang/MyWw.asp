<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
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
'Copyright (C) 2008��2009 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
keeper=Replace(Trim(Request.Form("keeper")),"'","''")
action=Replace(Trim(Request.Form("action")),"'","''")
if action="ok" then
  if Replace(Trim(request("code")),"'","''")<>czm then
  		 Response.Write("<script>alert('���������!');window.location.href='myww.asp';</script>")
		 response.End()
  end if
if not checkmaihao(keeper) then
  Response.Write("<script>alert('��󶨵ĵ����˺Ų�����Ч���˺�,����ȷ����!');window.location.href='myww.asp';</script>")
  response.End()
end if
		
Sql = "select count(*) as yu  from   "&jieducm&"_keeper where username='"&session("useridname")&"' and class=1"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF not(Rs.Eof) Then
coun=rs("yu")
end if

if jifei<2000 and coun>=2  then
  	Response.Write("<script>alert('2000�������£����ɰ�2������!');window.location.href='myww.asp';</script>")
	response.End()
elseif jifei>2000 and jifei<5000 and coun>=4 then
  	Response.Write("<script>alert('2000-5000���֣����ɰ�4������!');window.location.href='myww.asp';</script>")
	response.End()
elseif jifei>5000 and jifei<8000 and coun>=8 then
  	Response.Write("<script>alert('5000-8000���֣����ɰ�8������!');window.location.href='myww.asp';</script>")
	response.End()
elseif jifei>9000  and coun>=10 then
  	Response.Write("<script>alert('9000�������ϣ����ɰ�10������!');window.location.href='myww.asp';</script>")
	response.End()
end if

		  
      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_keeper where keeper='"&keeper&"'" ,Conn,3,3  
	  if rs.eof then
	    rs.addnew
		rs("username")=session("useridname")
		rs("keeper")=keeper
		rs("now")=now()
		rs("class")=1
    	rs.update
    	rs.close
		 Response.Write("<script>alert('�󶨳ɹ�!');window.location.href='myww.asp';</script>")
		 response.End()
	 else
	 	 Response.Write("<script>alert('���Ա������ƹ��ʺ��ѱ������û���!');history.back();</script>")
		 response.End()
	 end if
end if
%>	
<HTML><HEAD><title><%=stupname%> -<%=stuptitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<SCRIPT language=javascript>
function  save_onclick12()
{

  var Shopkeeper=document.aspnetForm.keeper.value;
  if(Shopkeeper=="")
  {
    alert("�������Ա������ƹ��ʺţ�");
	document.aspnetForm.keeper.focus();
	return false;
	}

 var code=document.aspnetForm.code.value;
  if(code=="")
  {
    alert("����������룡");
	document.aspnetForm.code.focus();
	return false;
	}

  return true;  
}
</SCRIPT>
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=../taobao/jieducm_toptao.asp-->
<div align="center">
<DIV 
style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 900px">
<UL id=task_nav>
  <LI><A  href="index.asp">�Ա��ղ���</A> </LI>
  <LI><A  href="pushmission.asp">���ղ�����</A> </LI>
  <LI><A  href="ReMission.asp">�ѽ�����</A> </LI>
  <LI><A  href="MyMission.asp">�ѷ�����</A> </LI>
    <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="MyWw.asp">�󶨵���</A> </LI>
  <LI><A href="Buyhao.asp">�����</A> </LI>
  <LI><A href="MySave.asp">����ֿ�</A> </LI>
  </UL></DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG 
src="../images/task_nav_line.gif"></DIV>
</DIV>
<DIV 
style="MARGIN-BOTTOM: 10px; WIDTH: 910px; PADDING-TOP: 1px; BACKGROUND-COLOR: #ffffff">
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV 
style="BACKGROUND: url(../images/task_bg_01.jpg) repeat-x 50% top; WIDTH: 910px; LINE-HEIGHT: 38px; HEIGHT: 38px; TEXT-ALIGN: center">���Ѱ󶨵��Ա������ʺţ��󶨵����ƹ�ź���ܷ����� 
</DIV></DIV>
<FORM id=aspnetForm name=aspnetForm  action=MyWw.asp method=post onSubmit="return save_onclick12()">
<DIV> </DIV>
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="WIDTH: 910px">
<TABLE 
style="BORDER-TOP-WIDTH: 0px; MARGIN-TOP: 10px; BORDER-LEFT-WIDTH: 0px; BORDER-BOTTOM-WIDTH: 0px; MARGIN-BOTTOM: 10px; WIDTH: 70%; BORDER-RIGHT-WIDTH: 0px" 
align=center>
  <TBODY>
  <TR>
    <TD align=right height=30>�Ա������ƹ��ʺţ�</TD>
    <TD><div align="left">
      <INPUT id=keeper style="WIDTH: 190px" maxLength=20 name=keeper>
    </div></TD></TR>
  <TR>
    <TD align=right height=30>�����룺</TD>
    <TD><div align="left">
      <INPUT id=code style="WIDTH: 190px"  type=password maxLength=20 name=code>
    </div></TD></TR>
  <TR>
    <TD height=30></TD>
    <TD><div align="left">
      <INPUT id=Button1 style="WIDTH: 200px; HEIGHT: 30px" type=submit value=���ҵĵ����ƹ��ʺ� name=Button1>
      <input type="hidden" name="action" value="ok">
    </div></TD></TR>
  <TR>
    <TD colSpan=2 height=50><SPAN 
      style="PADDING-LEFT: 120px; COLOR: red">ע��һ��ƽ̨�˺������԰�ʮ��û�������˺Ű󶨹��ĵ����ƹ������󶨺��޷��޸ģ�</SPAN></TD>
  </TR></TBODY></TABLE>
</DIV></DIV>
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="WIDTH: 910px">
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD>
      <TABLE cellSpacing=0 cellPadding=0 align=left border=0>
        <TBODY>
        <TR>
          <TD width=1><IMG height=25 src="../images/mytaobaobj1_3.gif" 
          width=3></TD>
          <TD align=middle width=180 
            background=../images/mytaobaobj1_4.gif><FONT 
            color=#ffffff>�Ѱ��Ա������ƹ������ʺ�</FONT></TD>
          <TD width=1><IMG height=25 src="../images/mytaobaobj1_6.gif" 
          width=3></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 align=left border=0>
        <TBODY>
        <TR>
          <TD>&nbsp;&nbsp;&nbsp; 2000�������£��ɰ�2�����̣�2000-5000���֣��ɰ�4�����̣�5000-9000���֣��ɰ�8�����̣�9000�������ϣ��ɰ�10������</TD>
        </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD bgColor=#E9A91E height=4></TD></TR></TBODY></TABLE></DIV></DIV>
<table width="69%" align="center" cellpadding="5" cellspacing="1" bgcolor="#D0DBE3" >
        <tr bgcolor="#E1F2FB">
          <td height="29" class="yell_font" style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">�Ա������ƹ� </div></td>
          <td class="yell_font" style="FONT-WEIGHT: bolder; WIDTH: 20%"><div align="center">��ʱ�� </div></td>
        </tr>
        <tbody>
		  <%
			  	Sql = "select * from "&jieducm&"_keeper  where username='"&session("useridname")&"' and class=1 order by id desc "
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("������Ϣ")
				Else
					Do While Not Rs.Eof
			  %>
          <tr>
            <td height="32" bgcolor="#FFFFFF" class=centers><div align="center"><%=rs("keeper")%> </div></td>
            <td bgcolor="#FFFFFF" class=centers><div align="center"><%=rs("now")%> </div></td>
          </tr>
		  	 <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
        </tbody>
      </table>


<DIV></DIV>

</FORM>
</DIV></div>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%rs.close
set rs=nothing
call closeconn()
%>
</BODY></HTML>
