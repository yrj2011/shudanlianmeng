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
function ConfirmDel()
{
	if(document.myform.Action.value=="Del")
	{
		document.myform.action="admin_recordmdel.asp";
		if(confirm("ȷ��Ҫɾ��ѡ�еļ�¼�𣿱����������޷��ָ���"))
		    return true;
		else
			return false;
	}
	else
	{
		document.myform.action="chongdel.asp";
		if(document.myform.TargetClassID.value=="")
		{
			alert("���ܽ������ƶ�����������Ŀ����Ŀ���ⲿ��Ŀ�У�");
			return false;
		}
		if(confirm("ȷ��Ҫ��ѡ�е������ƶ���ָ������Ŀ��"))
		    return true;
		else
			return false;
	}
}

</SCRIPT>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   

<DIV 
style="CLEAR: both; OVERFLOW: hidden; WIDTH: 950px; BACKGROUND-COLOR: #f3f8fe">
  <DIV 
style="CLEAR: both; MARGIN-BOTTOM: 10px; MARGIN-LEFT: 1px;MARGIN-TOP: 10px; COLOR: #ff0000">&nbsp;&nbsp;ѡ�����в�����¼��ʽ�� 
* <a href="admin_Recordm.asp">���в�����¼�б�</a> * <a href="admin_Recordm.asp?action=manage">��������б�</a> * <a href="admin_Recordm.asp?action=chong">��ֵ�����б�</a> * <a href="admin_Recordm.asp?action=login">��¼��̨�б�</a></DIV>
</DIV>
<table border=0 width=100% align=center bgcolor="#FFFFFF">
  <tr>
            <td width="100%"> 
              <div align="center">
                <form name="form1" method="post" action="admin_recordm.asp?action=ok">
                  
                  <IMG src="../../images/search.gif" width="16" height="16" align=absMiddle>��������Ա����������
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
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #e6f7fb">
<DIV style="CLEAR: both; WIDTH: 98%; LINE-HEIGHT: 35px; HEIGHT: 35px">
<UL class=mission>
  <LI style="WIDTH: 20%">�û��� </LI>
  <LI style="WIDTH: 10%">���� </LI>
  <LI style="WIDTH: 40%">��ϸ </LI>

  <LI style="WIDTH: 10%">��� </LI>
  <LI style="WIDTH: 18%">ʱ�� </LI></UL></DIV></DIV>
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid">
<form name="myform" method="Post" action="admin_recordmdel.asp" onSubmit="return ConfirmDel();">
<%	
action=request("action")

if action="" then
sql="select * from "&jieducm&"_recordm  order by id desc"
elseif action="chong" then
sql="select * from "&jieducm&"_recordm where   class='����' or class='��ֵ' or class='��ֵ' or class='����' or class='����' order by id desc"
elseif action="manage" then
sql="select * from "&jieducm&"_record where class='�������' or class='����ɾ��' or class='�����ϼ�' or class='ȷ�Ϸ���' or class='�򷽺���' or class='�������' order by id desc"
elseif action="login" then
sql="select * from "&jieducm&"_recordm where   class='��¼��̨' order by id desc"

elseif action="ok" then
sql="select * from "&jieducm&"_recordm where  username = '"&request("text")&"' or class like '%"&request("text")&"%' order by id desc"
end if

			Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then
					
	dim maxperpage,url,PageNo
	if action="" then
	   url="admin_recordm.asp"
	elseif action="ren" then
	   url="admin_recordm.asp?action=ren"
    elseif action="chong" then
	  	 url="admin_recordm.asp?action=chong"
    elseif action="zeng" then
	 	 url="admin_recordm.asp?action=zeng"
    elseif action="login" then
		 url="admin_recordm.asp?action=login"
   elseif action="manage" then
    	 url="admin_recordm.asp?action=manage"
   elseif action="ok" then
   	 url="admin_recordm.asp?action="&request("text")&""
  end if
	 if session("webid")<>"" then url=url&"&username="&session("webid")
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
	  %>
       <% DO WHILE NOT rs.EOF AND RowCount>0%>
<DIV 
style="WIDTH: 98%; PADDING-TOP: 10px; BORDER-BOTTOM: #06314a 1px dashed; HEIGHT: 50px">
<UL class=missionitem>
 
  <LI style="WIDTH: 20%"><input name='ID' type='checkbox' onClick="unselectall()" id="ID" style="border: 0px;background-color: #eeeeee;" value=<%=rs("id")%>><%=rs("username")%> </LI>
  <LI style="WIDTH: 10%"><%=rs("class")%> </LI>
  <LI style="WIDTH: 40%"><%=rs("content")%></LI>
  <LI style="WIDTH: 10%"><%=rs("jiegou")%> </LI>
  <LI style="WIDTH: 18%"><%=rs("now")%> </LI></UL>
</DIV>
   <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
<DIV 
style="WIDTH: 98%; LINE-HEIGHT: 40px; PADDING-TOP: 10px; HEIGHT: 40px; TEXT-ALIGN: center"> <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"������") %></DIV>
<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" style="border: 0px;background-color: #eeeeee;">
              ѡ�б�ҳ��ʾ�����м�¼
              <input name="submit" type='submit' value='&nbsp;ɾ��ѡ���ļ�¼&nbsp;' onClick="document.myform.Action.value='Del'"  style="cursor: hand;background-color: #cccccc;">
              <input name="Action" type="hidden" id="Action" value="Del"></form>
&nbsp;&nbsp;&nbsp;

</DIV>
<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>
