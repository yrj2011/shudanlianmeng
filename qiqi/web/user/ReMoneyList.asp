<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
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
<SCRIPT src="../js/Dialog.js" type=text/javascript></SCRIPT>
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
                    <div class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt; �����б� &gt;&gt; </div>
                    <div class=pp8><strong>�����б�</strong></div>
                    <br>
                    <br>
					  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
          <tr>
            <td width="125" height="26" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor border-bot"><div align="center">��ˮ��</div></td>
            <td width="142" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor  border-bot"><div align="center">���</div></td>
            <td width="143" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor  border-bot"><div align="center">���ֽ��պ�</div></td>
            <td width="143" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor  border-bot"><div align="center">����ʱ�� </div></td>
            <td width="94" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="red-bcolor  border-bot"><div align="center">״̬</div></td>
            <td width="111" align="center" background="images/luntan03.gif" bgcolor="#fffcd2" class="border-bot red-bcolor "><div align="center">����</div></td>
          </tr>
          <%	

   	sql="SELECT * FROM "&jieducm&"_tixian where username='"&session("useridname")&"' order by start desc"
  
			Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then
					
	dim maxperpage,url,PageNo
	 url="ReMoneyList.asp"
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
          <tr>
            <td height="35" align="center" class="border-botdashed"><div align="center"><%=rs("pid")%></div></td>
            <td align="center" class="border-botdashed"><div align="center"><%=rs("ReNum")%>Ԫ</div></td>
            <td align="center" class="border-botdashed"> &nbsp;
              <div align="center"><%=rs("ReZfb")%></div></td>
            <td align="center" class="border-botdashed">
              <div align="center"><%=rs("now")%></div></td>
            <td align="center" class="border-botdashed">			<div align="center">
              <%
			if(rs("start")=0) then
			   response.write("<font color=red>������</font>")
			elseif (rs("start")=1 )then
			   response.write("���ֳɹ�")
			 elseif (rs("start")=2 )then
			   response.write("��������")
			 elseif (rs("start")=3 ) then
			    response.write("������")
			 end if
			%>             
            </div></td>
            <td align="center" class="border-botdashed"><div align="center">
              <%	if(rs("start")=0) then%>
               <a href="#1" onClick="javascript:showDialog('��ȷ��Ҫ��������������',1,7,'tixiandel.asp?id=<%=rs("id")%>')" title="ȷ��Ҫ�����������֣�"><font color=red>��������</font></a>
               <%
			elseif (rs("start")=1 )then
			   response.write("��֧��")
			 elseif (rs("start")=2 )then
			   response.write("�ѳ���")
			    elseif (rs("start")=3 )then
			   response.write("������")
			 end if
			%> 
            </div></td>
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
<%call closeconn()%>
</BODY></HTML>
