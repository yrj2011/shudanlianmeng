<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
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
'Copyright (C) 2008��2011 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
id=request("id")
action=request("action")
if id<>"" and action="del" then
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_depot where id="&id&"  and  username='"&session("useridname")&"'",conn,3,3
response.Redirect("MySave.asp")	
elseif id<>"" and  action="ok" then
if cunkuan<1 then
 Response.Write("<script>alert('������1Ԫʱ�����ٷ�����');window.location.href='../user/md5_pay.asp';</script>")
 response.End()
elseif fabudian<1 then
 Response.Write("<script>alert('���������1��ʱ�����ٷ�����');window.location.href='../user/md5_pay.asp';</script>")
 response.End() 
end if	

sql="SELECT  count(*) as totalz FROM "&jieducm&"_zhongxin where classid='2' and username='"&session("useridname")&"' and start<>'5'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalzx=rs("totalz")
end if 

if int(totalzx)>=int(stupfaweiok) and vip<>"1" then
	Response.Write("<script>alert('������ͨ��Աֻ�ܷ���"&stupfaweiok&"��δ��ɵ�����\n��������ѷ�������!�����ټ���������');history.back();</script>")
    response.End()
elseif int(totalzx)>=int(stupfaweiok+5) and vip="1" then
	Response.Write("<script>alert('��VIP��Աֻ�ܷ���"&stupfaweiok+5&"��δ��ɵ�����\n��������ѷ�������!�����ټ���������');history.back();</script>")
    response.End()
end if 
			
   Set rs=server.createobject("ADODB.RECORDSET")
   rs.open "Select * From "&jieducm&"_depot where id="&id&" and  username='"&session("useridname")&"'",Conn,3,3
   if not rs.eof or not rs.bof then
    Price=rs("Price")
    ProUrl=rs("ProUrl")
    Shopkeeper=rs("Shopkeeper")
    bei=rs("bei")
    baohu2=rs("baohu2")
    isprice=rs("isprice")
    depot=rs("depot")
	isprice=rs("isprice")
	play=rs("play")
	tips=rs("tips")
    company=rs("company")
	numbers=rs("numbers")
	jieducm_fb=rs("jieducm_fb")
	else
	Response.Write("<script>alert('��������!');history.back();</script>")
	    response.End()
	end if
	


jieducm_ck=cunkuan*100
jieducm_price=price*100
fabudian2=fabudian*100
fabu2=jieducm_fb*100

if isfa="1" then
	Response.Write("<script>alert('���ѱ�����Ա��ֹ�˷�������!');history.back();</script>")
	response.End()
elseif(int(jieducm_price)>int(jieducm_ck)) then
 Response.Write("<script>alert('�㷢���������Ǯ���ܴ�����Ĵ��!');history.back();</script>")
 response.End()
elseif(int(fabudian2)<int(fabu2)) then
 Response.Write("<script>alert('��ķ�������������,�뾡���ֵ��ȡ!');history.back();</script>")
 response.End()
end if
	
		randomize
		ranNum=int(90000*rnd)+10000	
        pid=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
		Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_zhongxin where pid='"&pid&"'" ,Conn,3,3  
		if rs.eof then
	    rs.addnew
		rs("classid")=2
		rs("Price")=Price
    	rs("Shopkeeper")=Shopkeeper
		rs("ProUrl")=ProUrl
		rs("bei")=bei
		rs("baohu2")=baohu2
		rs("username")=session("useridname")
		rs("pid")=pid
		rs("now")=now
		rs("isprice")=isprice
		rs("play")=play
		rs("tips")=tips
		rs("company")=company
		rs("numbers")=numbers
		rs("jieducm_fb")=jieducm_fb
    	rs.update
    	rs.close
		else
		 Response.Write("<script>alert('������!�����·���');history.back();</script>")
         response.End()
		end if

      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
		rs("fabudian")=rs("fabudian")-jieducm_fb
    	rs("cunkuan")=rs("cunkuan")-price
    	rs.update
    	rs.close
				
sql="update "&jieducm&"_depot set num=num+1 ,now='"&now()&"'  where id="&id&""
conn.execute sql

     Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=pid
		rs("class")="����������"
		rs("content")=session("useridname")&"������ֿ��з�������"&pid&"�ɹ�,�����������"&jieducm_fb&"��"
		rs("price")="-"&price
		rs("jiegou")="�ɹ�"
    	rs.update
    	rs.close
		call check_jieducm_name(session("useridname")) 
	Response.Write("<script>alert('�����ɹ�!');window.location.href='MyMission.asp';</script>")
	response.End()
end if
%>
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<SCRIPT src="../js/Dialog.js" type=text/javascript></SCRIPT>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR></HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">    
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=../taobao/jieducm_toptao.asp-->

<div align="center">
<DIV 
style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 700px">
<UL id=paipai_task_nav>
  <LI><A  href="index.asp">���Ļ�����</A> </LI>
  <LI><A  href="pushmission.asp">��������</A> </LI>
  <LI><A href="javascript:getPart('ReMission1.asp',<%=request("pageno")%>,<%=request("sort")%>);">�ѽ�����</A> </LI>
  <LI><A href="MyMission.asp">�ѷ�����</A> </LI>
  <LI><A href="Buyhao.asp">�����</A> </LI>
   <LI><A href="MyWw.asp">�󶨵���</A> </LI>
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/paipai_task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff"  href="MySave.asp">����ֿ�</A> </LI></UL>
  </DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG src="../images/paipai_task_nav_line.gif"></DIV>
</DIV>
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 0px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 910px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #e6f7fb">
<DIV style="CLEAR: both; WIDTH: 98%; LINE-HEIGHT: 35px; HEIGHT: 35px">
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 125px; COLOR: #006600; TEXT-ALIGN: center">����ID</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 95px; COLOR: #006600; TEXT-ALIGN: center">�۸�</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 95px; COLOR: #006600; TEXT-ALIGN: center">��Ʒ��Ϣ</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 155px; COLOR: #006600; TEXT-ALIGN: center">��ע </DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 140px; COLOR: #006600; TEXT-ALIGN: center">�ϴη���ʱ�� 
</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 130px; COLOR: #006600; TEXT-ALIGN: center">�Ѿ�����</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 150px; COLOR: #006600; TEXT-ALIGN: center">��&nbsp;&nbsp;��</DIV></DIV></DIV>
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 910px; BORDER-BOTTOM: #abbec8 1px solid;BACKGROUND-COLOR: #ffffff">
 <%
	Sql = "select * from "&jieducm&"_depot where classid='2'  and username='"&session("useridname")&"' order by id desc"
   Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	if not rs.eof then
	url="mysave.asp"
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
	DO WHILE NOT rs.EOF AND RowCount>0
	  %>
			
<DIV class=missionbg 
style="WIDTH: 98%; PADDING-TOP: 10px; BORDER-BOTTOM: #06314a 1px dashed; HEIGHT: 70px">
<DIV style="FLOAT: left; WIDTH: 125px;LINE-HEIGHT: 150%;  TEXT-ALIGN: center">
<%=rs("title")%><br>
<%=rs("pid")%>  <br>
��ע��������<%
price=rs("price")
bei=rs("bei")
jieducm_fb=rs("jieducm_fb")
addfabu=0
call jieducm_fadian(jieducm_fb,addfabu)
 %>��<br>
   </DIV>
<DIV style="FLOAT: left; WIDTH: 95px; LINE-HEIGHT: 150%; TEXT-ALIGN: center"><font color=#ff6600>ƽ̨����<br>
      <img src="../images/j_z.gif" width="13" height="16"><%=rs("Price")%>Ԫ<br><%=rs("isprice")%></font></DIV>
<DIV style="FLOAT: left; WIDTH: 95px; LINE-HEIGHT: 150%; TEXT-ALIGN: center">
<span style="FLOAT: left;LINE-HEIGHT: 150%; WIDTH: 95px;  TEXT-ALIGN: center">
<input name="<%=rs("pid")%>" type="text" id="<%=pid%>" size="9" value="<%=rs("prourl")%>">
<INPUT style="WIDTH: 90px" onClick="jsCopy('<%=rs("pid")%>')" type=button value=������Ʒ��ַ>
 <BR>
�ƹ�<%=rs("Shopkeeper")%></span></DIV>
<DIV style="FLOAT: left; WIDTH: 155px; TEXT-ALIGN: center">
  <DIV style="CLEAR: both; WIDTH: 100%; HEIGHT: 20px"><%=rs("bei")%></DIV>
</DIV>
<DIV style="FLOAT: left; WIDTH: 140px; LINE-HEIGHT: 140%; TEXT-ALIGN: center"><%=rs("now")%></DIV>
<DIV style="FLOAT: left; WIDTH: 130px; LINE-HEIGHT: 120%; TEXT-ALIGN: center"><B 
style="COLOR: red"><%=rs("num")%> ��</B></DIV>
<DIV style="FLOAT: left; WIDTH: 150px; LINE-HEIGHT: 130%; TEXT-ALIGN: center">
<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫ������������������',1,7,'?id=<%=rs("id")%>&action=ok')" title="��ȷ��Ҫ������������������">
<SPAN class=anniu>��������</SPAN></A> 
<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫɾ������������',1,7,'?id=<%=rs("id")%>&action=del')" title="��ȷ��Ҫɾ������������">
<SPAN class=anniu>ɾ��</SPAN></A><BR><BR>
<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫ���±༭��������',1,7,'MySavePush.asp?id=<%=rs("id")%>')" title="��ȷ��Ҫ���±༭��������">
<SPAN class=anniu>���±༭</SPAN></A>
</DIV>
</DIV>
<%RowCount = RowCount - 1
i=i-1
rs.MoveNext 
Loop %>
<DIV 
style="WIDTH: 98%; LINE-HEIGHT: 40px; PADDING-TOP: 10px; HEIGHT: 40px; TEXT-ALIGN: center"> <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"������") %></DIV></DIV>
</div>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
call closeconn()
%>
</BODY></HTML>
