<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
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
<LINK href="../images/mission.css" type=text/css rel=stylesheet>
<LINK href="../images/Default.css" type=text/css rel=stylesheet>
<LINK href="../images/Bbs.css" type=text/css rel=stylesheet>
<LINK href="../images/style.css" type=text/css rel=stylesheet>
<SCRIPT src="../js/jieducm_Dialog.js" type=text/javascript></SCRIPT>

<title>�� �� �� ý</title>

<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   
<table border=0 width=100% align=center bgcolor="#FFFFFF">
  <tr>
            <td width="100%"> 
              <div align="center">
                <form name="form1" method="post" action="admin_mymission.asp?sort=ok&classid=1">
                  
                  <IMG src="../../images/search.gif" width="16" height="16" align=absMiddle>��������ID���û�����
                  <input class=input1 size=25 name=text value="<%=request.Form("text")%>">
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
<DIV style="CLEAR: both; OVERFLOW: hidden; WIDTH: 950px; BACKGROUND-COLOR: #f3f8fe"><IMG 
src="images/arrow2.gif">&nbsp;<SPAN 
style="FONT-WEIGHT: bolder; FONT-SIZE: 18px; COLOR: red">�ѷ��Ա�������</SPAN>&nbsp; 
* <A href="admin_MyMission.asp">��������</A>
* <A href="admin_MyMission.asp?sort=0">����������</A>
*<A href="admin_MyMission.asp?sort=1">�ȴ�����</A>
*<A href="admin_MyMission.asp?sort=2">�ȴ�֧��</A>
*<A href="admin_MyMission.asp?sort=3">�ȴ�����</A>
*<A href="admin_MyMission.asp?sort=4">�ȴ����ֺ���</A>
*<A href="admin_MyMission.asp?sort=5">�ȴ�����ȷ��</A>
*<A href="admin_MyMission.asp?sort=6">�Ѿ����</A>
</DIV>

<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #e6f7fb">
<DIV style="CLEAR: both; WIDTH: 98%; LINE-HEIGHT: 35px; HEIGHT: 35px">
<UL class=mission>
  <LI style="WIDTH: 15%">����ID </LI>
  <LI style="WIDTH: 10%">�۸� </LI>
  <LI style="WIDTH: 15%">��Ʒ��Ϣ </LI>
  <LI style="WIDTH: 15%">���ַ� </LI>
  <LI style="WIDTH: 15%">����ʱ�� </LI>
  <LI style="WIDTH: 15%">״&nbsp;&nbsp;̬</LI>
   <LI style="WIDTH: 15%">��&nbsp;&nbsp;��</LI>
  </UL></DIV></DIV>
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid">
<%
action=request.QueryString("sort")
classid=request("classid")
if action="" then
			  	Sql = "select * from "&jieducm&"_zhongxin where  classid='1'  order by start  asc"
elseif action="0" then 
			  	Sql = "select * from "&jieducm&"_zhongxin where  start<>'0' and start<>'5' and  classid='1' order by start  asc"
elseif action="1" then 
			  	Sql = "select * from "&jieducm&"_zhongxin where start='0'  and  classid='1' order by start  asc"
elseif action="2" then
			  	Sql = "select * from "&jieducm&"_zhongxin where start='1'   and  classid='1' order by start  asc"
elseif action="3" then
			  	Sql = "select * from "&jieducm&"_zhongxin where  start='2'  and  classid='1' order by start  asc"
elseif action="4" then
			  	Sql = "select * from "&jieducm&"_zhongxin where  start='3'   and  classid='1' order by start  asc"
elseif action="5" then
			  	Sql = "select * from "&jieducm&"_zhongxin where  start='4'  and  classid='1'  order by start  asc"
elseif action="6" then
			  	Sql = "select * from "&jieducm&"_zhongxin where  start='5'  and  classid='1'  order by start  asc"
elseif action="ok" then
 sql="select * from "&jieducm&"_zhongxin where  classid='1' and (pid like '%"&request("text")&"%' or  username like '%"&request("text")&"%')   order by start asc"				
end if
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
	dim maxperpage,url,PageNo
	if action="" then
	 url="admin_mymission.asp"
	elseif action="ok" then
	url="admin_mymission.asp?sort=ok&text="&request("text")&""
	else
	 url="admin_mymission.asp?sort="&action&""
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
 DO WHILE NOT rs.EOF AND RowCount>0
 pid=rs("pid")
 username=rs("username")
 start=rs("start")
 bei=rs("bei")
 id=rs("id")
 jieducm_fb=rs("jieducm_fb")
  addfabu=rs("addfabu")
  IfComeSet=rs("IfComeSet")
 %>
 <table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="100%">
<DIV style="WIDTH: 98%; PADDING-TOP: 10px; BORDER-BOTTOM: #06314a 1px dashed; HEIGHT: 100px">
<UL class=missionitem>
 
  <LI style="WIDTH: 15%">
  	<%if IfComeSet>0 then response.Write("<img src=../images/jieducm_lailu.gif border='0' title='��·��������'><br>") end if%>
  <a href="admin_recordu.asp?username=<%=username%>"><%=username%></a></font><br />
  ������:<font color="#FF0000"><%
  if addfabu>0 then
response.Write(jieducm_fb&"+"&addfabu)
else
response.Write(jieducm_fb)
end if
  %></font>��<br />
<%=pid%><br /><%=bei%> </LI>
  <LI style="WIDTH: 10%"><font color=#ff6600>ƽ̨����<br>
<img src="../images/jieducm_zf.gif" width="13" height="16"><%=rs("Price")%>Ԫ</font></LI>
  <LI style="WIDTH: 15%">
  <input name="<%=pid%>" type="text" id="<%=pid%>" size="9" value="<%=rs("prourl")%>">
<INPUT style="WIDTH: 90px" onClick="jsCopy('<%=pid%>')" type=button value=������Ʒ��ַ>

  <BR>
�ƹ�<%=rs("Shopkeeper")%>
<br />
 <%
			  	Sqlu = "select qq from "&jieducm&"_register where username='"&username&"'"
				Set Rsu = Server.CreateObject("Adodb.RecordSet")
				Rsu.Open Sqlu,conn,1,1
				if not rsu.eof then
				qq=rsu("qq")
				end if
				rsu.close
			  %>
<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=qq%>&site=qq&menu=yes"><img height=16 alt=���߻�ˢƽ̨��ӭ�� src="../images/qq1.gif" width=23 border=0>
</a>
</span></LI>
  <LI style="WIDTH: 15%">
  <%
			  	Sqlu = "select username,num,now from "&jieducm&"_jieshou where pid='"&pid&"'"
				Set Rsu = Server.CreateObject("Adodb.RecordSet")
				Rsu.Open Sqlu,conn,1,1
				if rsu.eof then
                response.Write("���޽�����")
				else
				nowj=rsu("now")
				num=rsu("num")
				usernamej=rsu("username")
				rsu.close
				end if
				
			  	Sqlu = "select username,qq from "&jieducm&"_register where username='"&usernamej&"'"
				Set Rsu = Server.CreateObject("Adodb.RecordSet")
				Rsu.Open Sqlu,conn,1,1
				if not rsu.eof then
				qqj=rsu("qq")
				rsu.close
				end if
			  %>
<%if start<>0 then%>			  
<a href="admin_recordu.asp?username=<%=usernamej%>"> <%=usernamej%></a><br>
<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=qqj%>&site=qq&menu=yes"><img height=16 alt=���߻�ˢƽ̨��ӭ�� src="../images/qq1.gif" width=23 border=0> </a>
<%end if%>
</LI>
  <LI style="WIDTH: 15%">
  <%if start=0 then%>
  ���޽�����
  <%else%>
  <%=nowj%>	<br />�����Ա��ţ�<font color="#FF0000"><%=num%></font>
  <%end if%>
   </LI>
  <LI style="WIDTH: 15%"><%if start=0 then%>
<B style="COLOR: red">�ȴ�������... </B>
<%elseif start=1 then%>
<B style="COLOR:#006600">�ѽ��֣�<BR>�ȴ����ַ�֧��...  </B>
<%elseif start=2 then%>
<B style="COLOR: #80b309">��֧����<BR>�ȴ�����������...</B>
<%elseif start=3 then%>
<B style="COLOR:blue">�ѷ�����<BR>�ȴ����ַ�����...   </B>
<%elseif start=4 then%>
<B style="COLOR: #000000">�Ѻ�����<BR>�ȴ�����������... </B>
<%elseif start=5 then%>
<B style="COLOR:#006600">��ɣ���ú��� </B>
<%end if%></LI>
   <LI style="WIDTH: 15%"><%if start="1" then%><a href="delmy.asp?action=qing&id=<%=pid%>" onClick="return confirm('ȷ��Ҫ������');">�������</a><%else%><font color="#999999">�������</font><%end if%>
|<%if start="0" then%><a href="delmy.asp?action=del&id=<%=id%>" onClick="return confirm('ȷ��Ҫɾ��������');">����ɾ��</a><%else%><font color="#999999">����ɾ��</font><%end if%>
<br>
<%if start="3" then%><a href="delmy.asp?action=hao&id=<%=pid%>">�򷽺���</a><%else%><font color="#999999">�򷽺���</font><%end if%>
|<%if start="4" then%><a href="delmy.asp?action=haook&id=<%=pid%>">�������</a><%else%><font color="#999999">�������</font><%end if%>
<br>
<%if start<>"5" and start<>"0"  and start<>"1" then%><a href="delmy.asp?action=back&id=<%=pid%>">�����ϼ�</a><%else%><font color="#999999">�����ϼ�</font><%end if%>
|<%if start="2" then%><a href="delmy.asp?action=ok&id=<%=pid%>">ȷ�Ϸ���</a><%else%><font color="#999999">ȷ�Ϸ���</font><%end if%></LI>
  </UL></DIV></td>
  </tr>
</table>
   <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
<DIV 
style="WIDTH: 98%; LINE-HEIGHT: 40px; PADDING-TOP: 10px; HEIGHT: 40px; TEXT-ALIGN: center"> 
<% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"������") %></DIV>

</DIV>
<%
set rsu=nothing
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>