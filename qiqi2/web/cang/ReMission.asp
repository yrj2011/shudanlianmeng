<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
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
<HTML lang=en><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<SCRIPT language=JavaScript src="../js/jquery.min.js"></SCRIPT>
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
   $.post('/cang/jieducm-ajax.asp?open=Reajex',function (data) {
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
<!--#include file=../taobao/jieducm_toptao.asp-->
<div align="center">
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 900px">
<UL id=task_nav>
  <LI><A  href="index.asp">�Ա��ղ���</A> </LI>
  <LI><A  href="pushmission.asp">���ղ�����</A> </LI>
  <LI><A style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="ReMission.asp">�ѽ�����</A> </LI>
  <LI><A  href="MyMission.asp">�ѷ�����</A> </LI>
    <LI><A  href="MyWw.asp">�󶨵���</A> </LI>
  <LI><A href="Buyhao.asp">�����</A> </LI>
  <LI><A href="MySave.asp">����ֿ�</A> </LI>
  </UL></DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG src="../images/task_nav_line.gif"></DIV>
</DIV>

<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="BACKGROUND: url(../images/task_bg_01.jpg) repeat-x 50% top; WIDTH: 910px; HEIGHT: 38px">
<DIV style="MARGIN-TOP: 10px; FLOAT: left"><IMG 
src="../images/task_02.gif"></DIV>
<DIV style="MARGIN-TOP: 12px; FLOAT: left; MARGIN-LEFT: 10px">
<A href="?sort=6"><SPAN class=anniu>�ҽ��ֽ���������</SPAN></A>
<A href="?sort=5"><SPAN class=anniu>�����</SPAN></A>
</DIV>
<DIV style="CLEAR: right; MARGIN-TOP: 12px; FLOAT: right"><A title=���ˢ�� href="javascript:void(0)" onClick="reloadTask()"><SPAN class=anniu id=shuaimg>ˢ��ҳ��  ���񲻶�����</SPAN></A></DIV></DIV></DIV>
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
Sql = "select * from "&jieducm&"_jieshou where classid='6'  and username='"&session("useridname")&"'  order by start asc ,id desc"
elseif action="6" then
Sql = "select * from "&jieducm&"_jieshou where classid='6'  and username='"&session("useridname")&"'  order by start asc"
else
Sql = "select * from "&jieducm&"_jieshou where classid='6'  and username='"&session("useridname")&"' and start='"&action&"' order by start  asc"
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
now1=rs("now")
prourl=rs("prourl")
Shopkeeper=rs("Shopkeeper")
num=rs("num")
start=rs("start")
id=rs("id")
Sql2 = "select * from "&jieducm&"_zhongxin where pid='"&pid&"'"
Set Rs2 = Server.CreateObject("Adodb.RecordSet")
Rs2.Open Sql2,conn,1,1
if not rs2.eof then
play=rs2("play")
nowf=rs2("now")
usernamef=rs2("username")
tips=rs2("tips")
bei=rs2("bei")
end if
 %>	  
<DIV class=missionbg style="WIDTH: 98%; PADDING-TOP: 10px; HEIGHT: 70px">
<DIV style="FLOAT: left; WIDTH: 125px; LINE-HEIGHT: 150%; TEXT-ALIGN: left"><%=pid%>
<BR><%=bei%>
<BR>
<%=nowf%>
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
</DIV>

</DIV>
<DIV style="FLOAT: left; WIDTH: 115px;  TEXT-ALIGN: center"><font color=#ff6600>ƽ̨����<br>
  <img src="../images/jieducm_zf.gif" width="13" height="16"><%=stupcang%>��������</font><br>
</DIV>
<DIV style="FLOAT: left; WIDTH: 115px; LINE-HEIGHT: 100%; TEXT-ALIGN: center">
    
  <label>
  <input name="<%=pid%>" type="hidden" id="<%=pid%>" size="9" value="<%=prourl%>">
  <INPUT style="WIDTH: 90px" onClick="jsCopy('<%=pid%>')" type=button value=������Ʒ��ַ>
  <br>
  </label>
  <a onClick="javascript:clipboardData.setData('text','<%=tips%>');alert('��ǩ�ѳɹ�����');return false;" href="#">�����Ʊ�ǩ��</a> </DIV>
<DIV style="FLOAT: left; WIDTH: 140px; LINE-HEIGHT: 150%; TEXT-ALIGN: center"><%=now1%><BR>
<font color="#FF0000">�����Ա��ţ�</font><br><%=num%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 120px; LINE-HEIGHT: 120%; TEXT-ALIGN: center">
<%if start=1 then
response.Write("<B style='COLOR: red'>�ȴ��ղ�...  </B>")
elseif start=2 then
response.Write("<B style='COLOR:#006600'>������ղ� </B>")
end if
%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 160px; LINE-HEIGHT: 120%; TEXT-ALIGN: center">
<%if start=1 then
sss= DateDiff( "n",now1, Now())
if 20-sss>=1 then
%>
<a href="#1" onClick="javascript:showDialog('��ȷ�����Ѿ��ղ�����',1,7,'jieducm_ok.asp?id=<%=pid%>')" >
<SPAN class=anniu>�Ѿ��ղ�</SPAN>
</A>
<BR> 
ʣ��<span style="color: #ff0000"><span id="min1"><%=20-sss%></span> </span>���ӿ��ղ�
<%else

sqlr="update "&jieducm&"_zhongxin set  jieshou_num=jieshou_num-1 where  pid='"&pid&"'  and classid='6'"
conn.execute sqlr

sqlrs="delete  from "&jieducm&"_jieshou where id='"&id&"'  and classid='6'"
conn.execute sqlrs

end if%>

<BR>��ϵ�������ɼ�ʱ��<BR>
<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫ�˳�����������',1,7,'jieducm_remis.asp?id=<%=id%>')" title="��ȷ��Ҫ�˳�����������">
<SPAN class=anniu>�˳�����</SPAN></A>
<%end if%>

</DIV>
</DIV>
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
</div></div>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</BODY></HTML>
