<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<LINK 
href="../images/mission.css" type=text/css rel=stylesheet>
<LINK 
href="../images/Default.css" 
type=text/css rel=stylesheet>
<LINK href="../images/Bbs.css" type=text/css rel=stylesheet>
<LINK href="../images/style.css" type=text/css rel=stylesheet>
<title>���ߴ�ý</title>
</head>

<body>
<div align="center"><p>
<%	 	
admincunkuan=0
adminfabudain=0
adminjifei=0
adminchina=0
adminalipay1=0
adminbuyfb=0
adminbuycar=0
sql="SELECT  sum(cunkuan) as admincunkuan FROM "&jieducm&"_chongjilu where class='��̨��ֵ' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
admincunkuan=rs("admincunkuan")
end if 

sql="SELECT  sum(fabudian) as adminfabudian FROM "&jieducm&"_chongjilu where class='��̨��ֵ' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
adminfabudian=rs("adminfabudian")
end if 

sql="SELECT  sum(jifei) as adminjifei FROM "&jieducm&"_chongjilu where class='��̨��ֵ' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
adminjifei=rs("adminjifei")
end if 

sql="SELECT  sum(cunkuan) as adminchina FROM "&jieducm&"_chongjilu where class='������ֵ' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
adminchina=rs("adminchina")
end if 

sql="SELECT  sum(cunkuan) as adminebao FROM "&jieducm&"_chongjilu where class='�ױ���ֵ' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
adminebao=rs("adminebao")
end if 

sql="SELECT  sum(cunkuan) as adminalipay FROM "&jieducm&"_chongjilu where class='֧������ֵ' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
adminalipay1=rs("adminalipay")
end if 

sql="SELECT  sum(fabudian) as adminbuyfb FROM "&jieducm&"_chongjilu where class='���򷢲���' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
adminbuyfb=rs("adminbuyfb")
end if
 
sql="SELECT  sum(cunkuan) as adminbuycar FROM "&jieducm&"_chongjilu where class='����㿨' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_adminbuycar=rs("adminbuycar")
end if 

sql="SELECT  sum(cunkuan) as adminbuycar2 FROM "&jieducm&"_chongjilu where class='��ֵ����ֵ' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_adminbuycar2=rs("adminbuycar2")
end if 


sql="SELECT  count(*) as dateren FROM "&jieducm&"_zhongxin where (datediff(day,[now],getdate())=0) and start='5'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
daterenxin=rs("dateren")
end if 


sql="SELECT  sum(cunkuan) as totalcunkuan FROM "&jieducm&"_chongjilu where class='��̨��ֵ' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalcunkuan=rs("totalcunkuan")
end if 

sql="SELECT  sum(fabudian) as totalfabudian FROM "&jieducm&"_chongjilu where class='��̨��ֵ' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalfabudian=rs("totalfabudian")
end if 

sql="SELECT  sum(jifei) as totaljifei FROM "&jieducm&"_chongjilu where class='��̨��ֵ' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totaljifei=rs("totaljifei")
end if 

sql="SELECT  sum(cunkuan) as totalchong FROM "&jieducm&"_chongjilu where class='��ֵ����ֵ' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalchongzhi=rs("totalchong")
end if 


sql="SELECT  sum(cunkuan) as totalchina FROM "&jieducm&"_chongjilu where class='������ֵ' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalchina=rs("totalchina")
end if 

sql="SELECT  sum(cunkuan) as totalebao FROM "&jieducm&"_chongjilu where class='�ױ���ֵ' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalebao=rs("totalebao")
end if 


sql="SELECT  sum(cunkuan) as totalalipay FROM "&jieducm&"_chongjilu where class='֧������ֵ' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalalipay=rs("totalalipay")
end if 

sql="SELECT  sum(fabudian) as totalbuyfb FROM "&jieducm&"_chongjilu where class='���򷢲���'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalbuyfb=rs("totalbuyfb")
end if

sql="SELECT  sum(cunkuan) as totalbuycar FROM "&jieducm&"_chongjilu where class='����㿨' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1

if not rs.eof then			
jieducm_totalbuycar=rs("totalbuycar")
end if

sql="SELECT  sum(cunkuan) as totalusermoney FROM "&jieducm&"_register "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalusermoney=rs("totalusermoney")
end if

sql="SELECT  sum(fabudian) as totaluserfabudian FROM "&jieducm&"_register "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totaluserfabudian=rs("totaluserfabudian")
end if


sql="SELECT  sum(price) as totalprice FROM "&jieducm&"_zhongxin where start<>'5' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalprice =rs("totalprice")
end if

sql="SELECT sum(ReNum) as totaltixian FROM "&jieducm&"_tixian  where start='1'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totaltixian =rs("totaltixian")
end if

sql="SELECT  sum(Money) as jieducm_alipay FROM AutoCZ where  Status=1 and (datediff(day,[EndDate],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_alipay_total=rs("jieducm_alipay")
end if

sql="SELECT  sum(Money) as jieducm_alipay_total FROM AutoCZ where   Status=1 "
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_alipay_totalok=rs("jieducm_alipay_total")
end if

rs.close
set rs=nothing%>              
 </p>
  <p>&nbsp;</p>
  <TABLE cellSpacing=1 cellPadding=1 width="451" align=center bgColor=#c2ddef 
border=0>
  <TBODY>
  <TR>
    <TD class=t1 vAlign=center align=middle width=130 bgColor=#ffffff 
      height=30><div align="right">���պ�̨��ֵ:</div></TD>
    <TD width="314" height=30 align=left vAlign=center bgColor=#ffffff>
      <div align="left">��<%=admincunkuan%>�������㣺<%=adminfabudian%>���֣�<%=adminjifei%></div></TD></TR>
  <TR>
    <TD class=t1 vAlign=center align=middle width=130 bgColor=#f6fcff 
      height=30><div align="right">����������ֵ:</div></TD>
    <TD vAlign=center align=left bgColor=#f6fcff height=30><%=adminchina%>Ԫ</TD></TR>
  <TR>
    <TD class=t1 vAlign=center align=middle width=130 bgColor=#ffffff 
      height=30><div align="right">����֧������ֵ:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=adminalipay1%>Ԫ</TD></TR>
  <TR>
    <TD class=t1 vAlign=center align=middle bgColor=#f6fcff 
      height=30><div align="right">����֧�����Զ���ֵ:</div></TD>
    <TD vAlign=center align=left bgColor=#f6fcff height=30><%=jieducm_alipay_total%>Ԫ</TD>
  </TR>
  <TR>
    <TD class=t1 vAlign=center align=middle bgColor=#f6fcff 
      height=30><div align="right">�����ױ���ֵ:</div></TD>
    <TD vAlign=center align=left bgColor=#f6fcff height=30><div align="left"><%=adminebao%>Ԫ</div></TD>
  </TR>
  <TR>
    <TD class=t1 vAlign=center align=middle width=130 bgColor=#f6fcff 
      height=30><div align="right">��������������:</div></TD>
    <TD vAlign=center align=left bgColor=#f6fcff height=30><%=adminbuyfb%>��</TD></TR>
  <TR>
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">���������㿨:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=jieducm_adminbuycar%>Ԫ</TD>
  </TR>
  <TR>
    <TD class=t1 vAlign=center align=middle width=130 bgColor=#ffffff 
      height=30><div align="right">���ճ�ֵ����ֵ:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=jieducm_adminbuycar2%>Ԫ</TD></TR>
  <TR style="mso-yfti-irow: 6">
    <TD class=t1 vAlign=center align=middle width=130 bgColor=#f6fcff 
      height=30><div align="right">��������ɺ���:</div></TD>
    <TD vAlign=center align=left bgColor=#f6fcff height=30><%=daterenxin%>��</TD></TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right"></div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30>&nbsp;</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">��̨��ֵ�ܶ�:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30>���:<%=totalcunkuan%> �����㣺<%=totalfabudian%> ���֣�<%=totaljifei%></TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">������ֵ�ܶ�:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=totalchina%>Ԫ</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">֧������ֵ�ܶ�:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=totalalipay%>Ԫ</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">֧�����Զ���ֵ�ܶ�:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=jieducm_alipay_totalok%>Ԫ</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">�ױ���ֵ�ܶ�:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=totalebao%>Ԫ</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">��ֵ����ֵ�ܶ</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=totalchongzhi%>Ԫ</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">�����������ܶ�:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=totalbuyfb%>��</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">�����㿨�ܶ�:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30>
	<%if not isnull(jieducm_totalbuycar) then
	 response.Write(formatnumber((jieducm_totalbuycar),2))
	 else
	 response.Write("0")
	end if%>
	Ԫ</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">ƽ̨��Ա�˻��ܶ�:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><div align="left">
	<%if not isnull(totalusermoney) then
	 response.Write(formatnumber((totalusermoney),2))
	else
	response.Write("0")
	end if%>
	Ԫ  
	�����㣺
	<%if not isnull(totaluserfabudian) then
	 response.Write(formatnumber((totaluserfabudian),2))
	else
	response.Write("0")
	end if%>
</div></TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle width=130 bgColor=#ffffff 
      height=30><div align="right">�����е������ܶ�:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30>
	<%if not isnull(totalprice) then
	 response.Write(formatnumber((totalprice),2))
	else
	response.Write("0")
	end if%>
	Ԫ</TD></TR>
  <TR style="mso-yfti-irow: 8">
    <TD class=t1 vAlign=center align=middle width=130 bgColor=#f6fcff 
      height=30><div align="right">��������ܶ�:</div></TD>
    <TD vAlign=center align=left bgColor=#f6fcff height=30>
	<%if not isnull(totaltixian) then
	 response.Write(formatnumber((totaltixian),2))
	else
	response.Write("0")
	end if%>
	Ԫ</TD></TR></TBODY></TABLE>
</div>
</body>
</html>
<%
conn.close
set conn=nothing
%>