<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷�����  �� �� ý
'�ٷ���ҳ��http://www.778892.com
'����֧�֣�xsh
'QQ;859680966
'�绰��15889679112
'��Ȩ����Ȩ�������ߴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2013 ���ߴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
did=request.Form("did")
action=request.Form("fa")
if action="ok" then
Price=Replace(Trim(Request.Form("Price1")),"'","''")
ProUrl=Request.Form("ProUrl1")
Shopkeeper=Request.Form("Shopkeeper")
bei=Request.Form("bei")
baohu2=request.Form("baohu2")
isprice=request.form("isprice")
play=request.form("play")
Limit=request.Form("Limit")
title=Replace(Trim(Request.Form("title")),"'","''")
tips=Replace(Trim(request("tips")),"'","''")
addfabu=request.Form("addfabu")
if price="" then
	Response.Write("<script>alert('��Ʒ����۲���Ϊ��!');history.back();</script>")
    response.End()
elseif int(price)<25  or int(price)>1000 then
	Response.Write("<script>alert('Ŀǰ��ƽֻ̨֧�ַ���25-1000Ԫ������ ������������!');history.back();</script>")
    response.End()
elseif Shopkeeper="" then
	Response.Write("<script>alert('�㻹û��ѡ���ƹ���!');history.back();</script>")
    response.End()	
elseif prourl="" then
	Response.Write("<script>alert('��Ʒ��ַ����Ϊ��!');history.back();</script>")
    response.End()
end if

url= ProUrl
wstr=getHTTPPage(url) 
Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a class=""hCard fn"" href=""[^""]+""[^>]*>(.*?)</a>"
set Matches=objRegExp.Execute(wstr)
jieducm_sk=Matches(0).SubMatches(0)

Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a class=""hCard fn"" title=""(.*?)"""
set Matches=objRegExp.Execute(wstr)
jieducm_sk3=Matches(0).SubMatches(0)


if price>=25 and price<=40 then
   fabu=1
elseif price>40 and price<=100 then
   fabu=2
elseif price>100 and price<=200 then
   fabu=3
elseif price>200 and price<=500 then
  fabu=4
elseif price>500 and price<=1000 then
  fabu=5
end if 

if bei="һ����ջ�����" then
fabu=fabu*2
elseif bei="������ջ�����"  then
fabu=fabu*2+1
elseif bei="������ջ�����"  then
fabu=fabu*2+2
end if
 

       Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_depot where id="&did&" and  username='"&session("useridname")&"'" ,Conn,3,3  
        if not rs.eof then
		rs("Price")=Price
    	rs("Shopkeeper")=Shopkeeper
		rs("ProUrl")=ProUrl
		rs("baohu")=baohu
		if baohu2="" then
		rs("baohu2")=0
		else
		rs("baohu2")=baohu2
		end if
		rs("now")=now
		rs("isprice")=isprice
		rs("play")=play
		rs("tips")=tips
		rs("Limit")=Limit
		rs("title")=title
		rs("addfabu")=addfabu
		rs("jieducm_fb")=fabu
    	rs.update
    	rs.close
		else
			Response.Write("<script>alert('��������!');history.back();</script>")
            response.End()
		end if
set rs=nothing
Response.Write("<script>alert('�޸ĳɹ���');window.location.href='MySave.asp';</script>")
response.End()	
conn.close
set conn=nothing
end if


call myww(1)

id=request("id")
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_depot where id="&id&" and  username='"&session("useridname")&"'",Conn,1,1
if not rs.eof or not rs.bof then
Shopkeeper=rs("Shopkeeper")
bei=rs("bei")
baohu2=rs("baohu2")
isprice=Replace(Trim(rs("isprice")),"'","''")
play=rs("play")
tips=rs("tips")
Price=rs("price")
ProUrl=rs("ProUrl")
Limit=rs("Limit")
title=rs("title")
addfabu=rs("addfabu")
else
    Response.Write("<script>alert('�����ˣ��뷵��!');history.back();</script>")
	response.End()
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh" lang="zh" dir="ltr">
<head>
<title>��Ա����-<%=stupname%> -<%=stuptitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="../css/style.css" type=text/css rel=stylesheet>
<LINK href="../css/base.css" type=text/css rel=stylesheet>
<LINK href="../css/layout.css" type=text/css rel=stylesheet>
</head>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=jieducm_toptao.asp-->
<div align="center">  
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 800px">
<UL id=task_nav>
 <li><a  href="index.asp">�Ա�������</a> </li>
  <LI><A  href="pushmission.asp">��������</A> </LI>
  <LI><A href="ReMission.asp">�ѽ�����</A> </LI>
  <LI><A href="MyMission.asp">�ѷ�����</A> </LI>
  <LI><A  href="MyWw.asp">�󶨵���</A> </LI>
  <LI><A href="Buyhao.asp">�����</A> </LI>
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="MySave.asp">����ֿ�</A> </LI>
   <li><a href="../user/Express.asp">���ɿ�ݵ�</a> </li></UL></DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG src="../images/task_nav_line.gif"></DIV>
</DIV>
</div>
  <table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>
        <table width="910"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            
            <td valign="top">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr height="1">
                  <td height="5"></td>
      </tr>
  </table>
              <DIV class=page>
<DIV class=addtask-wrap>
<DIV class=inner>
�޸�����ֿ���������֪�� 
  <UL>
  <LI>1. �����������±༭�������飻����۵�����ͷ����㡣
  </LI>
  </UL>
</DIV></DIV><br>
<DIV class=addtask-wrap>
<DIV class=inner>
  <TABLE cellSpacing=0 cellPadding=0 border=0>
  <TBODY>
  <FORM name=form  action=""  method=post>
  <INPUT type=hidden value=ok name=fa> 
   <input type="hidden" name="did" value="<%=id%>">
  <TR>
    <TH><div align="right">��Ʒ����ۣ�</div></TH>
    <TD><input name=Price1 id=Price1 size="10" maxlength=6   onKeyUp="if(isNaN(value))execCommand('undo')" value="<%=Price%>" >Ԫע��(�ü۸��ǰ����ʷѵ��ܼ۸�)1-40Ԫ����һ�������㣻41-100Ԫ�����������㣻101-200Ԫ�������������㣻201-500Ԫ�����ĸ������㣻501-1000Ԫ�����������</TD></TR>
  <TR>
    <TH><div align="right">��&nbsp; ��&nbsp; ����</div></TH>
    <TD > <%	Sql = "select * from "&jieducm&"_keeper  where username='"&session("useridname")&"'  and class=1 "
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	IF Rs.Eof Then
	 Response.Write("������Ϣ")
	Else
	Do While Not Rs.Eof
%><label><input  type='radio' <%if rs("keeper")=Shopkeeper then%> checked="checked" <%end if%> name='Shopkeeper'  id='Shopkeeper' value="<%=rs("keeper")%>"> <%=rs("keeper")%>
</label>
  <% Rs.MoveNext
   Loop
   End IF%></TD>
  </TR>
  <TR>
    <TH><div align="right">��Ʒ��ַ��</div></TH>
    <TD ><INPUT class=txtinput id=ProUrl1 style="WIDTH: 350px" maxLength=355 name=ProUrl1 value="<%=ProUrl%>"> 
    ����д��Ʒҳ����ȷ��ַ</TD></TR>
  
 
    <tr>
      <TH><div align="right">���ӷ�����(�ײ������ѡ)��</div></TH>
      <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><select name="addfabu">
                <option value="0" <%if addfabu=0 then%> selected="selected" <%end if%>>������</option>
                <option value="1" <%if addfabu=1 then%> selected="selected" <%end if%>>����1��</option>
                <option value="2" <%if addfabu=2 then%> selected="selected" <%end if%>>����2��</option>
                <option value="3" <%if addfabu=3 then%> selected="selected" <%end if%>>����3��</option>
                <option value="4" <%if addfabu=4 then%> selected="selected" <%end if%>>����4��</option>
                <option value="5" <%if addfabu=5 then%> selected="selected" <%end if%>>����5��</option>
              </select>         ÿ��һ������1�������㣬�緢�������ײͼ�������ɾ������</TD>
    </TR>
    <TH><div align="right">�۸��Ƿ���ȣ�</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><input name="isprice" type="radio" id="isprice" value="������" <%if isprice="������" then%> checked <%end if%>>
                              <span class="font12l"> ������</span> <input type="radio" name="isprice" id="isprice" value="���޸ļ۸�" <%if isprice="���޸ļ۸�" then%> checked <%end if%>>
                               <span class="font12l">���޸ļ۸�</span>&nbsp; ����۸�Ͱ����˷ѵ���Ʒ�ܼ۸��Ƿ���ȣ�</TD>
  </TR>
  <TR>
    <TH><div align="right">��̬���֣�</div></TH>
    <TD 
    style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
	            <input name="play" type="radio" value="ȫ����5��"  <%if play="ȫ����5��" then%> checked <%end if%>>
                              <span class="font12l">ȫ����5��</span> <input type="radio" name="play" value="ȫ�������"  <%if play="ȫ�������" then%> checked <%end if%>>
                             <span class="font12l"> ȫ�������</span></TD>
  </TR>
   <TR>
    <TH></TH>
    <TD>&nbsp;</TD>
   </TR>
  <TR>
  
  <TR>
    <TH><div align="right">���������ƣ�</div></TH>
    <TD><select name="Limit" >
                                <option value="1" <%if limit=1 then%> selected <%end if%>>������</option>
                                <option value="2" <%if limit=2 then%> selected <%end if%>>����100��������</option>
                                <option value="3" <%if limit=3 then%> selected <%end if%>>����300��������</option>
                                <option value="4" <%if limit=4 then%> selected <%end if%>>����ֻ����VIP��Ա</option>
                                                                                          </select></TD>
  </TR>
  <TR>
    <TH><div align="right">���������</div></TH>
    <TD><INPUT   name=tips class=txtinput id=tips style="WIDTH: 290px" maxlength="100" value="<%=tips%>">
    ��������ҿɼ���100������</TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL><INPUT id=baohu disabled type=checkbox CHECKED value=1 name=baohu> ����·���� </LABEL>
      <LABEL><INPUT id=baohu3 disabled type=checkbox CHECKED value=1 name=baohu3> �����걣��</LABEL>	  <input id=baohu32 disabled type=checkbox checked value=1 
      name=baohu32>
      �Զ���ⱦ����ַ���ƹ���</TD>
  </TR>
  
  <TR>
    <TH>&nbsp;</TH>
    <TD><input  name="baohu2" type="checkbox" id="baohu2" value="1" <%if baohu2=1 then%> checked="checked" <%end if%> />  
                  ���񱣻��������߽��������Ҫ����˲ſ��Կ�����Ʒ��ַ��</TD>
  </TR>
  <TR>
    <TH>&nbsp;</TH>
    <TD><LABEL></LABEL>&nbsp; �ֿⱸע(�������Լ��ֱ���Ʒ)�� 
                               <label>
                               <input name="title" type="text" maxlength="20" value="<%=title%>">
                               </label></TD></TR>
  
  <TR>
    <TD 
    style="PADDING-RIGHT: 0px; PADDING-LEFT: 200px; PADDING-BOTTOM: 5px; PADDING-TOP: 5px" 
    colSpan=2><input style="FONT-WEIGHT: bold; WIDTH: 120px; CURSOR: pointer; COLOR: #000000; HEIGHT: 26px" type="submit" name="button" id="button"  value="�����޸�" onClick="this.disabled=true;document.form.submit();"></TD></TR></FORM></TBODY></TABLE>
</DIV></DIV></DIV>    </td>
              </tr>
      </table>	    </td>
	    </tr>
  </table>
  <!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
  <%
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing%>
</body>
</html>