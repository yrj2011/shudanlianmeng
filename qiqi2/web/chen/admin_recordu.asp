<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
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
action=request("action")
username=request("username")
%>
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
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   

<DIV 
style="CLEAR: both; OVERFLOW: hidden; WIDTH: 950px; BACKGROUND-COLOR: #f3f8fe">
  <DIV 
style="CLEAR: both; MARGIN-BOTTOM: 10px; MARGIN-LEFT: 1px;MARGIN-TOP: 10px; COLOR: #ff0000">ѡ�����в�����¼��ʽ�� 
* <A href="admin_Recordu.asp?username=<%=username%>">���в�����¼�б�</A> 
* <A href="admin_Recordu.asp?username=<%=username%>&action=ren">��������б�</A>  
* <A href="admin_Recordu.asp?username=<%=username%>&action=chong">��ֵ�����б�</A>
* <A href="admin_Recordu.asp?username=<%=username%>&action=zeng">��ֵ�����б�</A>
* <A href="admin_Recordu.asp?username=<%=username%>&action=ti">���ֲ����б�</A>
* <A href="admin_Recordu.asp?username=<%=username%>&action=login">��¼�б�</A>
* <A href="admin_Recordu.asp?username=<%=username%>&action=manage">��������б�</A>
</DIV>
</DIV>
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #e6f7fb">
<DIV style="CLEAR: both; WIDTH: 98%; LINE-HEIGHT: 35px; HEIGHT: 35px">
<UL class=mission>
  <LI style="WIDTH: 20%">��ˮ�� </LI>
  <LI style="WIDTH: 10%">���� </LI>
  <LI style="WIDTH: 30%">��ϸ </LI>
  <LI style="WIDTH: 10%">��� </LI>
  <LI style="WIDTH: 10%">��� </LI>
  <LI style="WIDTH: 18%">ʱ�� </LI></UL></DIV></DIV>
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid">
<form name="myform" method="Post" action="admin_recorddel.asp" onSubmit="return ConfirmDel();">
<%	
if action="" then
sql="select * from "&jieducm&"_record where username='"&username&"' order by id desc"
elseif action="ren" then
sql="select * from "&jieducm&"_record where username='"&username&"' and (class='�Ա�������' or class='VIP������') order by id desc"
elseif action="chong" then
sql="select * from "&jieducm&"_record where username='"&username&"' and (class='��ֵ����ֵ' or class='������ֵ' or  class='֧������ֵ'  or class='�ױ���ֵ')  order by id desc"
elseif action="zeng" then
sql="select * from "&jieducm&"_record where username='"&username&"' and (class='����������' or class='���򷢲���'  or class='���ֻ�������'  or class='�����㻻���')  order by id desc"
elseif action="login" then
sql="select * from "&jieducm&"_record where username='"&username&"' and class='��¼��Ϣ'  order by id desc"
elseif action="ti" then
sql="select * from "&jieducm&"_record where username='"&username&"' and class='��������'  order by id desc"
elseif action="manage" then
sql="select * from "&jieducm&"_record where username='"&username&"' and (class='�������' or class='����ɾ��' or class='�����ϼ�' or class='ȷ�Ϸ���' or class='�򷽺���' or class='�������') order by id desc"
elseif action="ok" then
sql="select * from "&jieducm&"_record where username='"&username&"' or num like '%"&request("text")&"%' order by id desc"
end if
			Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then
				name1=rs("username")	
	dim maxperpage,url,PageNo
	if action="" then
	   url="admin_recordu.asp?username="&username&""
	elseif action="ren" then
	   url="admin_recordu.asp?username="&username&"&action=ren"
    elseif action="chong" then
	  	 url="admin_recordu.asp?username="&username&"&action=chong"
    elseif action="zeng" then
	 	 url="admin_recordu.asp?username="&username&"&action=zeng"
	elseif action="login" then
	 	 url="admin_recordu.asp?username="&username&"&action=login"
    elseif action="ti" then
		 url="admin_recordu.asp?username="&username&"&action=ti"
   elseif action="manage" then
    	 url="admin_recordu.asp?username="&username&"&action=manage"
   elseif action="ok" then
   	 url="admin_recordu.asp?username="&username&"&action="&request("text")&""
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
<DIV 
style="WIDTH: 98%; PADDING-TOP: 10px; BORDER-BOTTOM: #06314a 1px dashed; HEIGHT: 50px">
<UL class=missionitem>
 
  <LI style="WIDTH: 20%"><%=rs("num")%> </LI>
  <LI style="WIDTH: 10%"><%=rs("class")%> </LI>
  <LI style="WIDTH: 30%"><%=rs("content")%></LI>
  <LI style="WIDTH: 10%"><%=rs("price")%> </LI>
  <LI style="WIDTH: 10%"><%=rs("jiegou")%> </LI>
  <LI style="WIDTH: 18%"><%=rs("now")%> </LI></UL></DIV>
   <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
<DIV 
style="WIDTH: 98%; LINE-HEIGHT: 40px; PADDING-TOP: 10px; HEIGHT: 40px; TEXT-ALIGN: center"> <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"������") %></DIV>
<input name="Action" type="hidden" id="Action" value="Del"></form>
&nbsp;&nbsp;&nbsp;
<table border=0 width=100% align=center bgcolor="#FFFFFF">
  <tr>
            <td width="100%"> 
              <div align="center">
                <form name="form1" method="post" action="?action=ok&username=<%=request("username")%>">
                  
                  <IMG src="../../images/search.gif" width="16" height="16" align=absMiddle>������ˮ�ţ�
                  <input class=input1 size=25 name=text>
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
</DIV>
<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>
