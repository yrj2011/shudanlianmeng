<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<!--#INCLUDE FILE="../background.asp"-->

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
'Copyright (C) 2008��2009 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
 
%>
<HTML lang=en><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="../css/global.css" type=text/css rel=stylesheet>
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
   $.post('/cang/jieducm-ajax.asp?open=myajex',function (data) {
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
  <div style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
    <div style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
      <div style="FLOAT: left; OVERFLOW: hidden; WIDTH: 900px">
        <ul id=task_nav>
          <li><a  href="index.asp">�Ա��ղ���</a> </li>
          <li><a  href="pushmission.asp">���ղ�����</a> </li>
          <li><a href="ReMission.asp">�ѽ�����</a> </li>
          <li><a style="FONT-SIZE: 16px; BACKGROUND: url(../images/task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" href="MyMission.asp">�ѷ�����</a> </li>
		  <LI><A  href="MyWw.asp">�󶨵���</A> </LI>
          <li><a href="Buyhao.asp">�����</a> </li>
          <li><a href="MySave.asp">����ֿ�</a> </li>
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
<DIV style="MARGIN-TOP: 12px; FLOAT: left; MARGIN-LEFT: 10px">
<A href="?sort=0"><SPAN class=anniu>ȫ���ȴ��������� </SPAN></A>
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
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 120px; COLOR: #006600; TEXT-ALIGN: center">����ʱ��</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 135px; COLOR: #006600; TEXT-ALIGN: center">������ղ� </DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 130px; COLOR: #006600; TEXT-ALIGN: center">�ղؽ���</DIV>
<DIV style="FONT-WEIGHT: bold; FONT-SIZE: 14px; FLOAT: left; WIDTH: 150px; COLOR: #006600; TEXT-ALIGN: center">��&nbsp;&nbsp;��</DIV></DIV></DIV>
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 910px; BORDER-BOTTOM: #abbec8 1px solid;BACKGROUND-COLOR: #ffffff">
 <%
action=request.QueryString("sort")
if action="" then
Sql = "select * from "&jieducm&"_zhongxin where classid='6'  and username='"&session("useridname")&"'  order by start  asc ,id desc"
elseif action="6" then
Sql = "select * from "&jieducm&"_zhongxin where classid='6'  and username='"&session("useridname")&"'  order by start  asc ,now desc"
else
Sql = "select * from "&jieducm&"_zhongxin where classid='6'  and username='"&session("useridname")&"' and start='"&action&"' order by start  asc ,id desc"
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
bei=rs("bei")
start=rs("start")
prourl=rs("prourl")
Shopkeeper=rs("Shopkeeper")
nowf=rs("now")
num=rs("num")
jieshou_num=rs("jieshou_num")
tips=rs("tips")
 %> 			
<DIV class=missionbg style="WIDTH: 98%; PADDING-TOP: 10px; BORDER-BOTTOM: #06314a 1px dashed; HEIGHT: 110px">
<DIV style="FLOAT: left; WIDTH: 145px;LINE-HEIGHT: 150%;  TEXT-ALIGN: center">
<%=pid%>  <br>
<%=bei%><br>
�ղر�ǩ��<%=tips%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 115px; LINE-HEIGHT: 150%; TEXT-ALIGN: center">
<font color=#ff6600> 
<img src="../images/j_z.gif" width="13" height="16"><%=num*stupcang%>��������
</font>
</DIV>
<DIV style="FLOAT: left; WIDTH: 95px; LINE-HEIGHT: 150%; TEXT-ALIGN: center">
<input name="<%=pid%>" type="text" id="<%=pid%>" size="9" value="<%=prourl%>">
<INPUT style="WIDTH: 90px" onClick="jsCopy('<%=pid%>')" type=button value=������Ʒ��ַ>
 �ƹ�<%=Shopkeeper%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 120px; TEXT-ALIGN: center"><%=nowf%></DIV>
<DIV style="FLOAT: left; WIDTH: 135px; LINE-HEIGHT: 130%; TEXT-ALIGN: center">
<%
if jieshou_num=0 then
response.Write("������ɵ��ղ�")
else
Set rs_jieducm=server.createobject("ADODB.RECORDSET")
rs_jieducm.open "select count(*) as jieducm_numbers from "&jieducm&"_jieshou where pid='"&pid&"' and start='2' and classid='6'",conn,1,1
if  not rs_jieducm.eof then
jieducm_numbers=rs_jieducm("jieducm_numbers")
rs_jieducm.close
end if
response.Write("������ղ�����"&jieshou_num&"��")
end if
%>
</DIV>
<DIV style="FLOAT: left; WIDTH: 130px; LINE-HEIGHT: 120%; TEXT-ALIGN: center">
<STRONG><%=jieshou_num%>/<%=num%></STRONG> 
	<DIV class=speedjieducm><DIV class=rate style="WIDTH: <%=jieshou_num/num*100%>%; float: left;"></DIV></DIV>
</DIV>

<DIV style="FLOAT: left; WIDTH: 150px; LINE-HEIGHT: 130%; TEXT-ALIGN: center">
<%if jieshou_num=0 then%>
<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫȡ������������',1,7,'jieducm_del.asp?id=<%=pid%>')" title="��ʱ��û�˽��֣�ȡ�����&#13;ƽ̨���˻���ķ����㣬&#13;����ˢ�°�������ǰ��">
<SPAN class=anniu>ȡ������</SPAN></A><BR>
ȡ������ƽ̨���˻���ķ����㣬����ˢ�°�������ǰ��<br>
<a href="#1" onClick="javascript:showDialog('��ȷ��Ҫˢ�°�������ǰ��',1,7,'jieducm_shuaxin.asp?id=<%=pid%>')" title="��ʱ��û�˽���ˢ�°�������ǰ��">
<SPAN class=anniu>ˢ����ǰ</SPAN></A>
<%elseif jieshou_num=num then
response.Write("<B style='COLOR:#006600'>��ȫ����� </B>")
else
response.Write("<B style='COLOR: red'>���������...  </B>")
end if%>
</DIV>

<DIV style="WIDTH: 98%; LINE-HEIGHT: 20px; PADDING-TOP: 2px;  TEXT-ALIGN: center"></DIV>
</DIV>
<%RowCount = RowCount - 1
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
set conn=nothing%>
</BODY></HTML>
