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
<SCRIPT src="../js/jieducm_zDialog.js" type=text/javascript></SCRIPT>
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
   $.post('/taobao/jieducm-ajax.asp?open=Reajex',function (data) {
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

function jieducm_fun(iddstaobao)
{
    var taobaopid=iddstaobao;
	var diag = new Dialog();
	diag.Width = 600;
	diag.Height = 350;
	diag.Title = "�鿴��Ʒ����";
	diag.URL = "jieducm_look.asp?taobaoid="+taobaopid;
	diag.show();
}


</SCRIPT>
<!--#include file=../inc/jieducm_top.asp-->
<!--#include file=jieducm_toptao.asp-->
<div align="center">
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 800px">
<UL id=task_nav>
  <LI><A  href="index.asp">�Ա�������</A> </LI>
  <LI><A  href="pushmission.asp">��������</A> </LI>
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="javascript:void(0)" onClick="reloadTask()">�ѽ�����</A> </LI>
  <LI><A  href="MyMission.asp">�ѷ�����</A> </LI>
    <LI><A  href="MyWw.asp">�󶨵���</A> </LI>
  <LI><A href="Buyhao.asp">�����</A> </LI>
  <LI><A href="MySave.asp">����ֿ�</A> </LI>
   <li><a href="../user/Express.asp">���ɿ�ݵ�</a> </li>
  </UL></DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG src="../images/task_nav_line.gif"></DIV>
</DIV>

<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="BACKGROUND: url(../images/task_bg_01.jpg) repeat-x 50% top; WIDTH: 910px; HEIGHT: 38px">
<DIV style="MARGIN-TOP: 10px; FLOAT: left"><IMG 
src="../images/task_02.gif"></DIV>
<DIV style="MARGIN-TOP: 12px; FLOAT: left; MARGIN-LEFT: 10px"><A href="?sort=6"><SPAN class=anniu>�ҽ��ֽ���������</SPAN></A>
<A href="?sort=1"><SPAN class=anniu>�ȴ�֧��</SPAN></A>
 <A href="?sort=2"><SPAN class=anniu>�ȴ�����</SPAN></A>
  <A href="?sort=3"><SPAN class=anniu>�ȴ�����</SPAN></A>
  <A href="?sort=4"><SPAN class=anniu>�ȴ����</SPAN></A>
<A href="?sort=5"><SPAN class=anniu>�����</SPAN></A>
<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫ��ⳬ��12Сʱ��ȷ�Ϻ�����������\n������ִ��ʱ��ϳ��������ĵȴ���',1,7,'jieducm_checkstart.asp')" title="��ȷ��Ҫ��ⳬ��12Сʱ��ȷ�Ϻ�����������"><SPAN class=anniu>��ⳬ��12Сʱ��ȷ�Ϻ���������</SPAN></a>
</DIV>
<DIV style="CLEAR: right; MARGIN-TOP: 12px; FLOAT: right"><A title=���ˢ�� href="javascript:void(0)" onClick="reloadTask()"><SPAN class=anniu id=shuaimg>ˢ��ҳ�� 
���񲻶�����</SPAN></A></DIV></DIV></DIV>

<div style="width:512px;height:32px;z_index:5px; display:none; position:fixed; _position:absolute;" id="waitingimage"><img src="../images/jieducm_shua.gif"/></div>
<div id="content">
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 5px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 910px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #e6f7fb">
<DIV style="CLEAR: both; WIDTH: 98%; LINE-HEIGHT: 35px; HEIGHT: 35px">
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 125px; COLOR: #006600; TEXT-ALIGN: center">����ID</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 115px; COLOR: #006600; TEXT-ALIGN: center">������</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 115px; COLOR: #006600; TEXT-ALIGN: center">����۸�</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 115px; COLOR: #006600; TEXT-ALIGN: center">��Ʒ��Ϣ</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 140px; COLOR: #006600; TEXT-ALIGN: center">����ʱ�� 
</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 120px; COLOR: #006600; TEXT-ALIGN: center">״&nbsp;&nbsp;̬</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 160px; COLOR: #006600; TEXT-ALIGN: center">��&nbsp;&nbsp;��</DIV></DIV></DIV>
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 910px; BORDER-BOTTOM: #abbec8 1px solid;BACKGROUND-COLOR: #ffffff">
<%
action=request.QueryString("sort")
if action="" then
Sql = "select * from "&jieducm&"_jieshou where classid='1'  and username='"&session("useridname")&"'  order by start asc ,id desc"
elseif action="6" then
Sql = "select * from "&jieducm&"_jieshou where classid='1'  and username='"&session("useridname")&"'  order by start asc"
else
Sql = "select * from "&jieducm&"_jieshou where classid='1'  and username='"&session("useridname")&"' and start='"&action&"' order by start  asc"
end if
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if rs.eof then
 response.write("���޽�����Ϣ")
 end if
if not rs.eof then				
	if action="" then
	 url="remission.asp"
	else
	 url="remission.asp?sort="&action&""
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
now1=rs("now")
start=rs("start")
prourl=rs("prourl")
Shopkeeper=rs("Shopkeeper")
num=rs("num")
baohu2=rs("baohu2")

Sql2 = "select * from "&jieducm&"_zhongxin where pid='"&pid&"'"
Set Rs2 = Server.CreateObject("Adodb.RecordSet")
Rs2.Open Sql2,conn,1,1
if not rs2.eof then
play=rs2("play")
nowf=rs2("now")
usernamef=rs2("username")
tips=rs2("tips")
isprice=rs2("isprice")
addfabu=rs2("addfabu")
ping=rs2("ping")
pingtxt=rs2("pingtxt")
jieducm_fb=rs2("jieducm_fb")
IfComeSet=rs2("IfComeSet")
end if
 %>	  
<DIV class=missionbg style="WIDTH: 98%; PADDING-TOP: 10px; HEIGHT: 90px">
<DIV style="FLOAT: left; WIDTH: 125px; LINE-HEIGHT: 150%; TEXT-ALIGN: center">
<%if IfComeSet>0 then response.Write("<img src=../images/jieducm_lailu.gif border='0' title='��·��������'><br>") end if%>
<%=pid%>
<BR>��ע��<%if bei="�����ջ�����"  then
response.Write("<IMG alt=�����ջ����� src=../images/jieducm_xuni.gif>")
j=0
elseif bei="һ����ջ�����"  then
j=24
response.Write("<IMG alt=�ӳ��ջ� src=../images/jieducm_shiwu.gif>")
elseif bei="������ջ�����"  then
j=48
response.Write("<IMG alt=�ӳ��ջ� src=../images/jieducm_shiwu.gif>")
elseif bei="������ջ�����"  then
j=72
response.Write("<IMG alt=�ӳ��ջ� src=../images/jieducm_shiwu.gif>")
end if
response.Write("������")
call jieducm_fadian(jieducm_fb,addfabu)%>��<BR><%=nowf%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 115px; TEXT-ALIGN: center">
<DIV style="OVERFLOW: hidden; WIDTH: 100%;  TEXT-ALIGN: center">
<%	Sql = "select jifei,qq,vip from "&jieducm&"_register where username='"&usernamef&"'"
Set Rs1 = Server.CreateObject("Adodb.RecordSet")
Rs1.Open Sql,conn,1,1
if not rs1.eof then
jifei=rs1("jifei")
qq=rs1("qq")
vipf=Replace(Trim(rs1("vip")),"'","''")
rs1.close
end if
 %>
<%call isonline(usernamef)%>&nbsp;<a href="../user/send_message.asp?sendname=<%=usernamef%>" title="����վ����Ϣ" target="_blank"><%=usernamef%></a>
<%if vipf="1" then %><img src="../images/jieducm_vip.gif"  alt="���VIP,����ֵ��<%call vipxinyu(usernamef)%>"  border="0"/><% end if%>
<br>
<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=qq%>&site=qq&menu=yes">
<img src="http://wpa.qq.com/pa?p=1:<%=qq%>:41"  border="0"  />
</a>
<%call xinyu(jifei)%><br>
[<a href="../user/tousu.asp?pid=<%=pid%>&usernamet=<%=usernamef%>" target="_blank"><font color="#FF0000">����</font></a>]
[<a href="../user/name.asp?heiname=<%=usernamef%>" target="_blank"><font color="#FF0000">����</font></a>]<br />
[<a href="../user/sms.asp?fausername=<%=usernamef%>" target="_blank"><font color="#FF0000">���ֻ�����</font></a>]
</DIV>

</DIV>
<DIV style="FLOAT: left; WIDTH: 115px;  TEXT-ALIGN: center"><font color=#ff6600>ƽ̨����<br>
  <img src="../images/j_z.gif" width="13" height="16"><%=rs("Price")%>Ԫ</font><br><%=isprice%>
   
  </DIV>
<DIV style="FLOAT: left; WIDTH: 115px; LINE-HEIGHT: 100%; TEXT-ALIGN: center">
   <%if baohu2="2" or baohu2="0" then%>
	<%if IfComeSet>0 then%>
	<input  type="button" name="<%=pid%>" value="�鿴��Ʒ��ַ" onclick=jieducm_fun('<%=pid%>')><br>
	<%else%>
  <label>
  <input name="<%=pid%>" type="text" id="<%=pid%>" size="9" value="<%=prourl%>">
  <br>
  <INPUT style="WIDTH: 90px" onClick="jsCopy('<%=pid%>')" type=button value=������Ʒ��ַ>
  <br>
  </label><%end if%>
�ƹ�
<%
response.Write(Shopkeeper)
else
response.Write("�������ܵ���������QQ��ϵ����������˺���ܿ�����Ʒ��ַ")
end if%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 140px; LINE-HEIGHT: 150%; TEXT-ALIGN: center"><%=now1%><BR>
<font color="#FF0000">�����Ա��ţ�</font><br><%=num%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 120px; LINE-HEIGHT: 120%; TEXT-ALIGN: center">
<%if start=0 then%>
<B style="COLOR: red">�ȴ�������...  </B>
<%elseif start=1 then%>
<B style="COLOR:#006600">�ѽ��֣�<BR>�ȴ����ַ�֧��... </B>
<%elseif start=2 then%>
<B style="COLOR: #80b309">��֧����<BR>�ȴ�����������...  </B>
<%elseif start=3 then%>
<B style="COLOR:blue">�ѷ�����<BR>�ȴ����ַ�����...  </B>
<%elseif start=4 then%>
<B style="COLOR: #000000">�Ѻ�����<BR>�ȴ�����������... </B>
<%elseif start=5 then%>
<B style="COLOR:#006600">��ɣ���ú��� </B>
<%end if%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 160px; LINE-HEIGHT: 120%; TEXT-ALIGN: center">
<%if start=1 then
sss= DateDiff( "n",now1, Now())
if 20-sss>=1 then
if baohu2="1" then
response.Write("�������ܵ���������QQ��ϵ����������˺���ܿ�����Ʒ��ַ")
elseif baohu2="2"  or baohu2="0" then%>

<a href="#1" onClick="javascript:showDialog('��ȷ�����Ѿ�����������Ա�����֧������������\n������ܵ�ƽ̨������',1,7,'jieducm_zhifu.asp?id=<%=pid%>')" title="����20�����ڽ�����Ʒ��ַ��֧����֧����֧������ص���Ѿ�֧����">
<SPAN class=anniu>�Ѿ�֧��</SPAN>
</A>
<BR> 
ʣ��<span style="color: #ff0000"><span id="min1"><%=20-sss%></span> </span>���ӿ�֧��
<%end if%>

<BR>��ϵ�������ɼ�ʱ��<BR>
<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫ�˳�����������',1,7,'jieducm_remis.asp?id=<%=pid%>')" title="��ȷ��Ҫ�˳�����������">
<SPAN class=anniu>�˳�����</SPAN></A>
<%else%>
<SPAN class=anniu>
<%call jieducm_exitauto(pid,session("useridname"))%></SPAN><BR> 
��ϵ�������ɼ�ʱ��<BR>
<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫ�˳�����������',1,7,'jieducm_remis.asp?id=<%=pid%>')" title="��ȷ��Ҫ�˳�����������"><SPAN class=anniu>�˳�����</SPAN></A>
<%end if
elseif start=2 then%>
<a href="#1" onClick="javascript:showDialog('��ȷ�Ϸ���δ֧����\n���񽫷���δ֧��״̬��',1,7,'jieducm_weizhifu.asp?id=<%=pid%>')" title="����㻹δ֧�����뷵��δ֧��״̬��">
<SPAN class=anniu>��δ֧��</SPAN></A>
<BR>����㻹δ֧��
���� ��δ֧����
�ύ����ʵ��ID<BR>
<%elseif start=3 then
if j=0 then%>
<a href="#1" onClick="javascript:showDialog('��ȷ�������ѷ�������',1,7,'jieducm_haoping.asp?id=<%=pid%>')" title="��ȷ�������ѷ�������">
<SPAN class=anniu>���Ѻ���</SPAN></A>
<%else
nowj=rs("now")
sss= DateDiff( "h", nowj, Now())
if cint(j)-cint(sss)<1 then%>
<a href="#1" onClick="javascript:showDialog('��ȷ�������ѷ�������',1,7,'jieducm_haoping.asp?id=<%=pid%>')" title="��ȷ�������ѷ�������">
<SPAN class=anniu>���Ѻ���</SPAN></A>
<%else
response.Write("���������ջ�<br>��ʣ")
response.Write(cint(j)-cint(sss))
response.Write("Сʱ")
end if
end if
elseif start=4 then%>
����֪ͨ������ȷ��<br>���������������
<%end if%></DIV>
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
<a onClick="javascript:clipboardData.setData('text','<%=pingtxt%>');alert('�Զ��������ѳɹ�����');return false;" href="javascript:void(0)"><Font color="#339933"><strong>�������Զ��������</strong></Font></a>
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
set conn=nothing
%>
</BODY></HTML>
