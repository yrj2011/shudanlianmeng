<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->
<!--#include file="../background.asp"-->
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ**************************************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷����ݶȴ�ý
'�ٷ���ҳ��http://www.jieducm.cn
'����֧�֣�robin 
'QQ;616591415
'�绰��13598006137
'��Ȩ����Ȩ���ڽݶȴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2008��2011 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------

%>
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="../css/center.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   
<!--#include file=../inc/jieducm_top.asp-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<tr>
	    <td>
        <table width="910"  border="0" cellspacing="0" cellpadding="0">
	          <tr>
	          
	            <td width="140" valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"> </td>
    </tr>
</table>



<table id="menuToolBar" width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
    <td valign="top" bgcolor="FFFFFF">
    <!--#include file="userleft.asp"--></TD>
</TR>
</table></td>
	            <td width="15">&nbsp;</td>
	          
	            <td valign="top">
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"></td>
    </tr>
</table>
                <div class=pp9>
                  <div style="PADDING-BOTTOM: 15px; WIDTH: 97%">
                    <div class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt; ������¼ &gt;&gt; </div>
                    <div class=pp8><strong>*<a href="Record.asp">���в�����¼�б�</a> *<a href="Record.asp?action=ren">��������б�</a> * <a href="Record.asp?action=chong">��ֵ�б�</a> * <a href="Record.asp?action=chongmanage">��̨��ֵ�б�</a> * <a href="Record.asp?action=zeng">��ֵ�����б�</a> * <a href="Record.asp?action=ti">���ֲ����б�</a> * <a href="Record.asp?action=manage">��������б�</a></strong></div>
                   
                  </DIV>
                  <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" >
          
          <tr>
            <td width="125" height="26" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor border-bot">��ˮ��</td>
            <td width="142" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor  border-bot">����</td>
            <td width="143" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor  border-bot">��ϸ</td>
            <td width="80" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor  border-bot">���</td>
            <td width="94" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor  border-bot">���</td>
            <td width="111" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="border-bot red-bcolor ">ʱ��</td>
          </tr>
          <%	
		  action=request("action")
		  if action="ren" then
            sql="select * from "&jieducm&"_record where username='"&session("useridname")&"' and (class='�Ա�������' or class='VIP������' or class='����������' or class='�а�������') order by id desc"
		  elseif action="chong" then
             sql="SELECT * FROM "&jieducm&"_record where   username='"&session("useridname")&"' and (class='��ֵ����ֵ' or class='������ֵ' or class='֧������ֵ') order by id desc"
		  elseif action="chongmanage" then
             sql="SELECT * FROM "&jieducm&"_record where   username='"&session("useridname")&"' and (class='����' or class='��ֵ' or class='��ֵ' or class='����' or class='����') order by id desc"
		 elseif action="zeng" then
sql="select * from "&jieducm&"_record where username='"&session("useridname")&"' and (class='����������' or class='���򷢲���'  or class='���ֻ�������'  or class='�����㻻���')  order by id desc"
         elseif action="ti" then
            sql="select * from "&jieducm&"_record where username='"&session("useridname")&"' and class='��������'  order by id desc"
		elseif action="manage" then
             sql="select * from "&jieducm&"_record where username='"&session("useridname")&"' and  (class='�������' or class='����ɾ��' or class='�����ϼ�' or class='ȷ�Ϸ���' or class='�򷽺���' or class='�������' or class='�ͷ���¼' or class='����VIP���û�' or class='�����˳�VIP��') order by id desc"
		  else  
      	        sql="SELECT * FROM "&jieducm&"_record where username='"&session("useridname")&"' order by id desc"
         end if
			Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then
					
	dim maxperpage,url,PageNo
	if action="" then
	   url="record.asp"
	else
	   url="record.asp?action="&action&""
  end if
	 rs.pagesize=10
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
          <tr>
            <td height="35" align="center" class="border-botdashed"><%=rs("num")%></td>
            <td align="center" class="border-botdashed"><%=rs("class")%></td>
            <td align="center" class="border-botdashed"><%=rs("content")%></td>
            <td align="center" class="border-botdashed"> &nbsp;<%=rs("price")%></td>
            <td align="center" class="border-botdashed">&nbsp; <%=rs("jiegou")%></td>
            <td align="center" class="border-botdashed"><%=rs("now")%> </td>
          </tr>
             <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
          <tr>
            <td height="35" colspan="6" align="center" class="border-botdashed"><% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"������") %></td>
            </tr>
        </table>
                  </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
call closeconn()%>
</BODY></HTML>
