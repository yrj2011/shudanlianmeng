<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
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
<HTML lang=en><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<SCRIPT src="../js/Dialog.js" type=text/javascript></SCRIPT>
<META content="MSHTML 6.00.2900.5921" name=GENERATOR></HEAD>
<BODY>
<SCRIPT language=JavaScript type=text/javascript>
 var time =  new Date();
   function reloadTask()
   {
   now = new Date();
   if((now-time)<3000)
   {
   alert('���3��,��Դ����,�벻Ҫ�ظ�Ƶ��ˢ��');
   }
   else
   {
   document.all["shuaimg"].innerHTML="���ݶ�ȡ��......";
   time = new Date();
   waitingimagePosition();
   $('#waitingimage').show();
   param = location.search.substring(1);
   if(param!=''){
param ='&'+	param;	   
   }
   $.post('/taobao/jieducm-ajax.asp?open=myajex',function (data) {
$('#content').html(data);
   $('#waitingimage').hide();
    document.all["shuaimg"].innerHTML="ˢ��ҳ��   ���񲻶�����";
   });
   //���ó�ʱ���Զ��ر������ ��ֹ���ֳ�ʱ�䲻���� �û���Ϊ��ס����� �����رպ����û���������û������µ��һ��
   setTimeout(function(){$('#waitingimage').hide();}, 1000);
   }
   }
 
function waitingimagePosition(){
  		var obj = $("#content");
 	var offset = obj.offset();
 	var middle = offset.left+440;
 	var down = offset.top+200; 
$("#waitingimage").css({
"position": "absolute",
"top": down,
"left": middle
});
  	}
</SCRIPT>
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=jieducm_toptao.asp-->

<div align="center">
  <div style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
    <div style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
      <div style="FLOAT: left; OVERFLOW: hidden; WIDTH: 800px">
        <ul id=task_nav>
          <li><a  href="index.asp">�Ա�������</a> </li>
          <li><a  href="pushmission.asp">��������</a> </li>
          <li><a href="ReMission.asp">�ѽ�����</a> </li>
          <li><a style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="javascript:void(0)" onClick="reloadTask()">�ѷ�����</a> </li>
		  <LI><A  href="MyWw.asp">�󶨵���</A> </LI>
          <li><a href="Buyhao.asp">�����</a> </li>
          <li><a href="MySave.asp">����ֿ�</a> </li>
		   <li><a href="../user/Express.asp">���ɿ�ݵ�</a> </li>
        </ul>
      </div>
    </div>
    <div style="CLEAR: both; WIDTH: 910px"><img 
src="../images/task_nav_line.gif"></div>
  </div>
  <DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="BACKGROUND: url(images/task_bg_01.jpg) repeat-x 50% top; WIDTH: 910px; HEIGHT: 38px">
<DIV style="MARGIN-TOP: 10px; FLOAT: left"><IMG 
src="../images/task_02.gif"></DIV>
<DIV style="MARGIN-TOP: 12px; FLOAT: left; MARGIN-LEFT: 10px"><A href="?sort=0"><SPAN class=anniu>ȫ���ȴ���������</SPAN></A>
<A href="?sort=1"><SPAN class=anniu>�ѽ���</SPAN></A>
 <A href="?sort=2"><SPAN class=anniu>��֧��</SPAN></A>
  <A href="?sort=3"><SPAN class=anniu>�ѷ���</SPAN></A>
  <A href="?sort=4"><SPAN class=anniu>�Ѻ���</SPAN></A>
<A href="?sort=5"><SPAN class=anniu>�����</SPAN></A></DIV>
<DIV style="CLEAR: right; MARGIN-TOP: 12px; FLOAT: right"><A title=���ˢ�� href="javascript:void(0)" onClick="reloadTask()"><SPAN class=anniu id=shuaimg>ˢ��ҳ�� 
���񲻶�����</SPAN></A></DIV></DIV></DIV>

<div style="width:512px;height:32px;z_index:5px; display:none; position:fixed; _position:absolute;" id="waitingimage"><img src="../images/jieducm_shua.gif"/></div>
<div id="content">
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 0px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 910px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #e6f7fb">
<DIV style="CLEAR: both; WIDTH: 98%; LINE-HEIGHT: 35px; HEIGHT: 35px">
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 145px; COLOR: #006600; TEXT-ALIGN: center">����ID</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 115px; COLOR: #006600; TEXT-ALIGN: center">�۸�</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 95px; COLOR: #006600; TEXT-ALIGN: center">��Ʒ��Ϣ </DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 120px; COLOR: #006600; TEXT-ALIGN: center">���ַ�</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 135px; COLOR: #006600; TEXT-ALIGN: center">����ʱ�� </DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 130px; COLOR: #006600; TEXT-ALIGN: center">״&nbsp;&nbsp;̬</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 150px; COLOR: #006600; TEXT-ALIGN: center">��&nbsp;&nbsp;��</DIV></DIV></DIV>
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 910px; BORDER-BOTTOM: #abbec8 1px solid;BACKGROUND-COLOR: #ffffff">
<%
action=request.QueryString("sort")
if action="" then
Sql = "select * from "&jieducm&"_zhongxin where classid='1'  and username='"&session("useridname")&"'  order by start  asc ,id desc"
elseif action="6" then
Sql = "select * from "&jieducm&"_zhongxin where classid='1'  and username='"&session("useridname")&"'  order by start  asc ,now desc"
else
Sql = "select * from "&jieducm&"_zhongxin where classid='1'  and username='"&session("useridname")&"' and start='"&action&"' order by start  asc ,id desc"
end if
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
	if action="" then
	url="mymission.asp"
	else
	 url="mymission.asp?sort="&action&""
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
 DO WHILE NOT rs.EOF AND RowCount>0
 pid=rs("pid")
price=rs("price")
bei=rs("bei")
start=rs("start")
prourl=rs("prourl")
Shopkeeper=rs("Shopkeeper")
play=rs("play")
isprice=rs("isprice")
nowf=rs("now")
addfabu=rs("addfabu")
ping=rs("ping")
pingtxt=rs("pingtxt")
tips=rs("tips")
jieducm_fb=rs("jieducm_fb")
IfComeSet=rs("IfComeSet")
 %> 	  
<DIV class=missionbg style="WIDTH: 98%; PADDING-TOP: 10px; HEIGHT: 90px">
<DIV style="FLOAT: left; WIDTH: 145px;LINE-HEIGHT: 150%;  TEXT-ALIGN: center">
<%if IfComeSet>0 then response.Write("<img src=../images/jieducm_lailu.gif border='0' title='��·��������'><br>") end if%>
<%=pid%>  <br>
��ע��<%
if bei="�����ջ�����"  then
response.Write("<IMG alt=�����ջ����� src=../images/jieducm_xuni.gif>")
j=0
elseif bei="һ����ջ�����"  then
j=24
response.Write("<IMG alt=�ӳ��ջ� src=../images/jieducm_shiwu.gif>")
elseif bei="������ջ�����"  then
j=48
response.Write("<IMG alt=�ӳ��ջ� src=../images/jieducm_shiwu.gif>")
elseif bei="������ջ�����" then
j=72
response.Write("<IMG alt=�ӳ��ջ� src=../images/jieducm_shiwu.gif>")
end if
response.Write("������")
call jieducm_fadian(jieducm_fb,addfabu)%>��
<br><%=nowf%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 115px; TEXT-ALIGN: center">
<DIV style="FLOAT: left; WIDTH: 115px; LINE-HEIGHT: 150%; TEXT-ALIGN: center">
<font color=#ff6600>ƽ̨����<br>
<img src="../images/j_z.gif" width="13" height="16"><%=Price%>Ԫ<br><%=isprice%></font>
</DIV>

</DIV>
<DIV style="FLOAT: left; WIDTH: 95px; LINE-HEIGHT: 150%; TEXT-ALIGN: center">
<input name="<%=pid%>" type="text" id="<%=pid%>" size="9" value="<%=prourl%>">
<INPUT style="WIDTH: 90px" onClick="jsCopy('<%=pid%>')" type=button value=������Ʒ��ַ>
 �ƹ�<%=Shopkeeper%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 120px; TEXT-ALIGN: center">
<%Sql2 = "select * from "&jieducm&"_jieshou where pid='"&pid&"'"
Set Rs2 = Server.CreateObject("Adodb.RecordSet")
Rs2.Open Sql2,conn,1,1
if rs2.eof then	
jie=0	
response.Write("���޽�����")

else
baohu2=rs2("baohu2")
jie=1
nowj=rs2("now")
numj=rs2("num")
ipj=rs2("ip")

Sql1 = "select * from "&jieducm&"_register where username='"&rs2("username")&"'"
Set Rs1 = Server.CreateObject("Adodb.RecordSet")
Rs1.Open Sql1,conn,1,1
if not rs1.eof then
jifei=rs1("jifei") 
username=rs1("username")
%>
<%call isonline(username)%>&nbsp;<a href="../user/send_message.asp?sendname=<%=username%>" title="����վ����Ϣ" target="_blank"><%=username%></a>
 <br>
<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=rs1("qq")%>&site=qq&menu=yes">
<img src="http://wpa.qq.com/pa?p=1:<%=rs1("qq")%>:41"  border="0"  />
</a>
<%
call xinyu(jifei)
if vip="1" then
response.Write("<br>���ַ�IP:")
response.Write(ipj)
end if
%>
<br><a href="../user/tousu.asp?pid=<%=pid%>&usernamet=<%=username%>" target="_blank"><font color="#FF0000">
[����]</font></a>
<a href="../user/name.asp?heiname=<%=username%>" target="_blank"><font color="#FF0000">
[����]</font></a>
<br />[<a href="../user/sms.asp?fausername=<%=username%>" target="_blank"><font color="#FF0000">���ֻ�����</font></a>]
<%
rs1.close
End IF
rs2.close
end if %>
</DIV>
<DIV style="FLOAT: left; WIDTH: 135px; LINE-HEIGHT: 130%; TEXT-ALIGN: center">

<%if jie=0 then
response.Write("���޽�����")	
else
response.Write(nowj)
response.Write("<br><font color=#FF0000>�����Ա��ţ�</font>")
response.Write(numj)
end if
%>

</DIV>
<DIV style="FLOAT: left; WIDTH: 130px; LINE-HEIGHT: 120%; TEXT-ALIGN: center">
<%if start=0 then%>
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
<%end if%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 150px; LINE-HEIGHT: 130%; TEXT-ALIGN: center">
<%if start=0 then%>
<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫȡ������������',1,7,'jieducm_del.asp?id=<%=pid%>')" title="��ʱ��û�˽��֣�ȡ�����&#13;ƽ̨���˻���ĵ�����ͷ����㣬&#13;����ˢ�°�������ǰ��">
<SPAN class=anniu>ȡ������</SPAN></A><BR>
ȡ������ƽ̨���˻���ĵ�����ͷ����㣬����ˢ�°�������ǰ��<br>
<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫˢ�°�������ǰ��',1,7,'jieducm_shuaxin.asp?id=<%=pid%>')" title="��ʱ��û�˽���ˢ�°�������ǰ��">
<SPAN class=anniu>ˢ����ǰ</SPAN></A>
<%elseif start=1 then 
sss= DateDiff( "n",nowj, Now())
if 20-sss<1 then
call jieducm_exitauto(pid,username)
else%>

<%if baohu2="1" then%>
<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫ���ͨ���ý�������\n��˺���ַ����Կ�����Ʒ��ַ��',1,7,'jieducm_shehe.asp?id=<%=pid%>')" title="��˺���ַ����Կ�����Ʒ��ַ��">
<SPAN class=anniu>��˽�����</SPAN></A><br>
<%if vip=1 then%>

<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫ���������\n���񽫷���δ����״̬��',1,7,'jieducm_qingli.asp?id=<%=pid%>')" title="������ң����񷵻�δ����״̬">
<SPAN class=anniu>�������</SPAN></A><br>
<%end if%>
<%else%>
<%if vip=1 then%>

<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫ���������\n���񽫷���δ����״̬��',1,7,'jieducm_qingli.asp?id=<%=pid%>')" title="������ң����񷵻�δ����״̬">
<SPAN class=anniu>�������</SPAN></A><br>
<%end if%>
�����<br>
<%end if%>
ʣ��<span style="color: #ff0000"><span id="min1"><%=20-sss%></span> </span>�����ӿ�֧��
<%end if%>
<br>
<a href="#1" onClick="javascript:showDialog('��ȷ��ҪΪ�Է���ʱ��',1,7,'jieducm_jiashi.asp?id=<%=pid%>')" title="Ϊ�Է���ʱ��">
<SPAN class=anniu>Ϊ�Է���ʱ</SPAN></A>
<%elseif start=2 then%>
<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫ�����ѷ���������\n����˶�֧�����Ƿ�֧���ɹ�,�������ѷ�������',1,7,'jieducm_fahou.asp?id=<%=pid%>')" title="����˶�֧�����Ƿ�֧���ɹ�,�������ѷ�������">
<SPAN class=anniu>���ѷ���</SPAN></A><BR>��ʵ���뷢����ƽ̨����
<%elseif start=3 then
if j=0 then
response.Write("��ϵ���ַ�������")
else
sss= DateDiff( "h", nowj, Now())
if cint(j)-cint(sss)<1 then
response.Write("��ϵ���ַ�������<br>�ջ�ʱ���ѵ�")
else
response.Write("���ַ����������ջ�<br>��ʣ")
response.Write(cint(j)-cint(sss))
response.Write("Сʱ")
end if
end if
elseif start=4 then%>
<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫȷ�Ϻ�����',1,7,'jieducm_ok.asp?id=<%=pid%>')" title="����˶�֧�����Ƿ�֧���ɹ�,�������ѷ�������!">
<SPAN class=anniu>ȷ�Ϻ���</SPAN></A><BR>������ȷ�Ϻ�������������ɣ�
<%end if%>
</DIV>
</DIV>
<DIV style="WIDTH: 98%; LINE-HEIGHT: 10px; PADDING-TOP: 10px;  TEXT-ALIGN: center">  
<div align="left">���̶�̬��֣�<%=play%>!<%=ping%>! <%=bei%>!<font color="#FF0000"> ����������ϵ </font></div> 
</DIV>

 <%if tips<>"" then%>
 <DIV style="WIDTH: 98%; LINE-HEIGHT: 10px; PADDING-TOP: 10px;  TEXT-ALIGN: center">  
 <div align="left">  
 ������ʾ��<font color="#FF0000"><%=tips%></font>
 </div></DIV>
 <%end if%>
  
<%if ping="�Զ�������" then%>
<DIV style="WIDTH: 98%; LINE-HEIGHT: 20px; PADDING-TOP: 2px;  TEXT-ALIGN: center">  
<div align="left">  �Զ������<font color="#FF0000"><%=pingtxt%></font>
</div>
</DIV>
<%end if%> 

<DIV style="WIDTH: 98%; LINE-HEIGHT: 3px; BORDER-BOTTOM: #06314a 1px dashed;">
</DIV>   
<%
RowCount = RowCount - 1
i=i-1
rs.MoveNext 
Loop
%>
<DIV style="WIDTH: 98%; LINE-HEIGHT: 40px; PADDING-TOP: 10px; HEIGHT: 40px; TEXT-ALIGN: center">
 <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"������") %></DIV></DIV>
</div>

</div>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>
