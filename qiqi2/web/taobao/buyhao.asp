<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="checklogin.asp"-->
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷����λ�����
'�ٷ���ҳ��http://www.x8888888.com
'�������ƽ̨��http://www.x8888888.com
'��ʾƽ̨��ַ��http://www.x8888888.com
'����֧�֣�robin 
'QQ;2227551
'�绰��15295958361
'��Ȩ����Ȩ�����λ���Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2008��2018 �λ���Ϣ�������޹�˾ ��Ȩ����
'ע: �벻Ҫ���Ը��Ĵ˱�ʶ,����Ҫ˽�·���������,������Ϊ�Զ������ۺ����
'*****************************************************************
'------------------------------------------------------------------

action=request("action")
if action="ok" then
jieducm_hao=Replace(Trim(Request.Form("hao")),"'","''")

Sql = "select count(*) as xinyu from "&jieducm&"_xinyu where username='"&session("useridname")&"' and class=1 and start=0 "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF not Rs.Eof Then
 ii=rs("xinyu")
End IF
if ii>=stupbuyhao then
 Response.Write("<script>alert('һ����ֻ�ܰ�"&stupbuyhao&"�����!');history.back();</script>")
response.End()
end if
 
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_xinyu where shopname='"&jieducm_hao&"' and class=1",Conn,3,3  
if rs.eof then
rs.addnew
rs("username")=session("useridname")
rs("shopname")=jieducm_hao
rs("num")=10
rs("class")=1	
rs("start")=0
rs.update
rs.close
Response.Write("<script>alert('�󶨳ɹ�!');window.location.href='buyhao.asp';</script>")
response.End()
else
 Response.Write("<script>alert('�˺��ѱ������û���!');history.back();</script>")
response.End()
end if
elseif action="del" then
id=request.QueryString("id")
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_xinyu where id="&id&" and username='"&session("useridname")&"'",conn,3,3
rs.close
Response.Write("<script>window.location.href='buyhao.asp';</script>")
response.End()
end if
%>	
<HTML><HEAD><title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<SCRIPT src="../js/Dialog.js" type=text/javascript></SCRIPT>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=jieducm_toptao.asp-->
<SCRIPT language=javascript>
function  save_onclick12()
{

  var Shopkeeper=document.form1.hao.value;
  if(Shopkeeper=="")
  {
    alert("�������ҵ�ˢ����ţ�");
	document.form1.hao.focus();
	return false;
	}

  return true;  
}
</SCRIPT>
<div align="center">
<div 
style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
    <div style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
      <div style="FLOAT: left; OVERFLOW: hidden; WIDTH: 800px">
        <ul id=task_nav>
          <li><a  href="index.asp">�Ա�������</a> </li>
          <li><a  href="<%if vip="1" then%>pushmissionvip.asp<%else%>pushmission.asp<%end if%>">��������</a> </li>
          <li><a href="ReMission.asp">�ѽ�����</a> </li>
          <li><a  href="MyMission.asp">�ѷ�����</a> </li>
		  <LI><A  href="MyWw.asp">�󶨵���</A> </LI>
          <li><a style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="Buyhao.asp">�����</a> </li>
          <li><a href="MySave.asp">����ֿ�</a> </li>
		   <li><a href="../user/Express.asp">���ɿ�ݵ�</a> </li>
        </ul>
      </div>
    </div>
    <div style="CLEAR: both; WIDTH: 910px"><img 
src="../images/task_nav_line.gif"></div>
  </div>
<DIV 
style="MARGIN-BOTTOM: 10px; WIDTH: 910px; PADDING-TOP: 15px; BACKGROUND-COLOR: #ffffff">

 <div class="pub_tip4">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td valign="bottom" width="120" align="center"><strong class="f_b_org"><font color="#EC8C11">��ʾ˵��</font></strong></td>
        <td><div align="left">1.��ҳ����Ҫ�����󶨺�ά�����������񣬹���������Ʒ���Ա���ţ�<a href="/news.asp?/1415.html" target="_blank" class="link_o">ʲô����ţ�</a><br />
          2.�����Ա������°�ȫ���ԣ�������Ŷ�Ҫ�������������д������Ϣ����ܽ��а󶨣�Ϊ�����������Է����갲ȫ�����Ƚ������Ϣ���ƣ�<br />
          3.Ϊ�˱�֤��������Լ�����������İ�ȫ���뾡���󶨶����ѡ��ţ���������ʱ�ֻ�ʹ�ã�<a href="/news.asp?/1480.html" target="_blank" class="link_o">ΪʲôҪ�󶨶����ţ�</a><br />
          4.Ϊ�˱�֤��������Լ�����������İ�ȫ��ƽ̨����ÿ���Ա����ÿ�ս�����������Ҫ����6����ÿ�ܲ�Ҫ����35��!<br />
          5.Ϊ�˱�֤ˢ�������������Ŷȣ�ϵͳΪÿ��������Զ�����һ��������ֵ�������������������ȴﵽ������ֵ���󽫲��ܽ����µ�����<br /></div></td>
      </tr>
    </table>   
  </div>
<DIV style="BACKGROUND: #ffffff; WIDTH: 900px">
<FORM id=form1 name=form1  action="" onSubmit="return save_onclick12()" method=post>
<input name="action" type="hidden" value="ok">
<DIV></DIV>
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="WIDTH: 910px">
<TABLE width="96%" 
align=center 
style="BORDER-TOP-WIDTH: 0px; MARGIN-TOP: 10px; BORDER-LEFT-WIDTH: 0px; BORDER-BOTTOM-WIDTH: 0px; MARGIN-BOTTOM: 10px; WIDTH: 50%; BORDER-RIGHT-WIDTH: 0px">
  <TBODY>
  <TR>
    <TD align=right height=30>�Ա������ʺţ�</TD>
    <TD><div align="left">
      <INPUT id=hao style="WIDTH: 190px" maxLength=20 name=hao>
    </div></TD></TR>
  
  <TR>
    <TD height=30></TD>
    <TD><div align="left">
      <INPUT id=Button1 style="WIDTH: 200px; HEIGHT: 30px"  type=submit value=���ҵ��Ա������ʺ� name=Button1>
    </div></TD></TR>
  <TR>
    <TD height=50></TD>
    <TD>&nbsp;</TD>
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
            color=#ffffff>�Ѱ��Ա������ʺ�</FONT></TD>
          <TD width=1><IMG height=25 src="../images/mytaobaobj1_6.gif" 
          width=3></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 align=left border=0>
        <TBODY>
        </TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD bgColor=#E9A91E height=4></TD></TR></TBODY></TABLE></DIV></DIV>
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="WIDTH: 910px">
<TABLE width="100%" align="center" cellpadding="5" cellspacing="1" bgcolor="#D0DBE3" >
        <TR bgcolor="#E1F2FB">
          <TD height="30" style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">�Ա������ʺ� </div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">��ʼ�������� </div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">����Ѿ����� </div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">�ú�����껹�����</div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 15%"><div align="center">״̬ </div></TD>
          <TD style="FONT-WEIGHT: bolder; WIDTH: 5%"><div align="center">���� </div></TD>
        </TR>
        <TBODY>
		  <%
			 
			
			  	Sql = "select * from "&jieducm&"_xinyu  where username='"&session("useridname")&"' and class=1  order by id desc "
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("������Ϣ")
				Else
					Do While Not Rs.Eof
			  %>
          <TR>
              <TD bgcolor="#FFFFFF" class=centers><img src="../images/j_z.gif" width="13" height="16"><%=rs("shopname")%></TD>
    <TD bgcolor="#FFFFFF" class=centers><%=rs("num")%></TD>
    <TD bgcolor="#FFFFFF" class=centers><%
	Sqlj = "select count(*) as jxinyu from "&jieducm&"_jieshou where username='"&session("useridname")&"' and classid='1' and num='"&rs("shopname")&"' "
	Set Rsj = Server.CreateObject("Adodb.RecordSet")
	Rsj.Open Sqlj,conn,1,1
	IF not Rsj.Eof Then
	 i=rsj("jxinyu")
	 else
	 i=0
	End IF
	response.Write(i)
	%></TD>
    <TD bgcolor="#FFFFFF" class=centers><%=251-(rs("num")+i)%>
                    <%xiaoyu=(rs("num")+i)
				  if xiaoyu>=251 then
				    sqlr="update "&jieducm&"_xinyu set start=1 where id="&rs("id")&""
                    conn.execute sqlr
				  end if
				  %></TD>
<TD bgcolor="#FFFFFF" class=centers><%if rs("start")="0" then%>���Լ���ʹ�� <%else%> �Ѵﵽ����<%end if%></TD>
    <TD bgcolor="#FFFFFF" class=centers><a href="#1" onClick="javascript:showDialog('��ȷ��Ҫɾ���������',1,7,'?action=del&id=<%=rs("id")%>')" title="��ȷ��Ҫɾ���������">ɾ��</A></TD>
          </TR>
		    <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
        </TBODY>
      </TABLE>
</DIV></DIV>
<DIV> </DIV>
</FORM>
</DIV>
</DIV></div>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%rs.close
set rs=nothing
call closeconn()
%>
</BODY></HTML>
