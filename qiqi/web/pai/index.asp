<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
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
 
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select top 1 num from "&jieducm&"_jieshou where username='"&session("useridname")&"' and classid='2'  order by  id desc",conn,1,1
if not (rs.eof) then
  checknum=rs("num")
rs.close
end if
%>
<HTML lang=en><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
 <META content="MSHTML 6.00.2900.5921" name=GENERATOR></HEAD>
<BODY>
<SCRIPT language=JavaScript type=text/javascript>

function showxiao(id,stat){
var idxiao=id;
sAlert("<div style=margin:10px;line-height:40px;><FORM id=xiaohao method=post><div style=line-height:40px;>  ���Ѱ󶨵����ĺš�<br><%	
			  	Sql = "select * from "&jieducm&"_xinyu  where username='"&session("useridname")&"' and class=2 and start=0 order by id desc "
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
				    wei=0
					Response.Write("<a href=buyhao.asp>�Բ����㻹û�а���ţ�����ź󣬲ſ��Խ�������!</a>")
				Else
				    wei=1
					Do While Not Rs.Eof
			  %> <input type=radio name=xiaohaoname <%if checknum=rs("shopname") then%> checked <%end if%> value=<%=rs("shopname")%> /><%=rs("shopname")%>&nbsp;&nbsp; <%
			  	Rs.MoveNext
				Loop
				End IF
			  %><br><%if wei=1 then %><INPUT type=submit value=ȷ��>&nbsp;<input type=button value=�ر� id=\"do_OK\" onclick=\"doOk()\" /><%end if%><%if wei=0 then%><input type=button value=�ر� id=\"do_OK\" onclick=\"doOk1()\" /><% end if%></div><div style=line-height:40px;></div></FORM></div>");
document.getElementById('xiaohao').action='jieducm_jieshou.asp?id='+idxiao;
}

function sAlert(txt){
    var eSrc=(document.all)?window.event.srcElement:arguments[1];
    var shield = document.createElement("DIV");
    shield.id = "shield";
    shield.style.position = "absolute";
    shield.style.left = "0px";
    shield.style.top = "0px";
    shield.style.width = "100%";
    shield.style.height = document.body.scrollHeight+"px";
    shield.style.background = "#333";
    shield.style.textAlign = "center";
    shield.style.zIndex = "10000";
    shield.style.filter = "alpha(opacity=0)";
    shield.style.opacity = 0;
    var alertFram = document.createElement("DIV");
    alertFram.id="alertFram";
    alertFram.style.position = "absolute";
    alertFram.style.left = "50%";
    alertFram.style.top = "50%";
    alertFram.style.marginLeft = "-225px" ;
    alertFram.style.marginTop = -75+document.documentElement.scrollTop+"px";
    alertFram.style.width = "500px";
    alertFram.style.height = "100px";
    alertFram.style.background = "#ccc";
    alertFram.style.textAlign = "center";
    alertFram.style.lineHeight = "100px";
    alertFram.style.zIndex = "10001";
    strHtml = "<ul style=\"list-style:none;margin:0px;padding:0px;width:100%\">\n";
    strHtml += "    <li style=\"background:#DD828D;text-align:left;padding-left:20px;font-size:14px;font-weight: bold;color:#fff;height:25px;line-height:25px;border:1px solid #F9CADE;\">[��ѡ���㹺�����Ʒ�������û�����]</li>\n";
    strHtml += "    <li style=\"background:#fff;text-align:center;font-size:12px;height:120px;line-height:80px;border-left:1px solid #F9CADE;border-right:1px solid #F9CADE;\">"+txt+"</li>\n";
    strHtml += "    <li style=\"background:#FDEEF4;text-align:left;font-weight:bold;height:25px;line-height:25px; border:1px solid #F9CADE;\"><span style='color: #0000FF;'>�ٷ���ʾ�����ϸ�������ʾ�����Ҫ�������������ӣ�Υ���涨�߷�ţ�</span><br\>\n";
    strHtml +="";
    strHtml +="</li>\n";
    
    strHtml += "</ul>\n";
    alertFram.innerHTML = strHtml;
    document.body.appendChild(alertFram);
    document.body.appendChild(shield);
    this.setOpacity = function(obj,opacity){
        if(opacity>=1)opacity=opacity/100;
        try{ obj.style.opacity=opacity; }catch(e){}
        try{ 
            if(obj.filters.length>0&&obj.filters("alpha")){
                obj.filters("alpha").opacity=opacity*100;
            }else{
                obj.style.filter="alpha(opacity=\""+(opacity*100)+"\")";
            }
        }catch(e){}
    }
    var c = 0;
    this.doAlpha = function(){
        if (++c > 20){clearInterval(ad);return 0;}
        setOpacity(shield,c);
    }
    var ad = setInterval("doAlpha()",1);
    this.doOk = function(){
        //alertFram.style.display = "none";
        //shield.style.display = "none";
        document.body.removeChild(alertFram);
        document.body.removeChild(shield);
        eSrc.focus();
        document.body.onselectstart = function(){return true;}
        document.body.oncontextmenu = function(){return true;}
    }
	    this.doOk1 = function(){
        //alertFram.style.display = "none";
        //shield.style.display = "none";
        document.body.removeChild(alertFram);
        document.body.removeChild(shield);
        eSrc.focus();
        document.body.onselectstart = function(){return true;}
        document.body.oncontextmenu = function(){return true;}
		window.location.href='buyhao.asp'
    }
    document.getElementById("do_OK").focus();
    eSrc.blur();
    document.body.onselectstart = function(){return false;}
    document.body.oncontextmenu = function(){return false;}
}
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
   $.post('/pai/jieducm-ajax.asp?open=indexajex',function (data) {
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

<DIV 
style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; OVERFLOW: hidden; WIDTH: 910px">
<DIV style="CLEAR: both; MARGIN-TOP: 2px; WIDTH: 910px; HEIGHT: 25px">
<DIV style="FLOAT: left; OVERFLOW: hidden; WIDTH: 700px">
<UL id=paipai_task_nav>
  <LI><A 
  style="FONT-SIZE: 16px; BACKGROUND: url(../images/paipai_task_nav_02.gif) no-repeat 50% bottom; COLOR: #ffffff" 
 href="javascript:void(0)" onClick="reloadTask()">���Ļ�����</A> </LI>
  <LI><A  href="pushmission.asp">��������</A> </LI>
  <LI><A href="ReMission.asp">�ѽ�����</A> </LI>
  <LI><A href="MyMission.asp">�ѷ�����</A> </LI>
   <LI><A href="MyWw.asp">�󶨵���</A> </LI>
  <LI><A href="Buyhao.asp">�����</A> </LI>
  <LI><A href="MySave.asp">����ֿ�</A> </LI></UL>
  </DIV>
</DIV>
<DIV style="CLEAR: both; WIDTH: 910px"><IMG 
src="../images/paipai_task_nav_line.gif"></DIV>
</DIV>

<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV 
style="BACKGROUND: url(../images/task_bg_01.jpg) repeat-x 50% top; WIDTH: 910px; HEIGHT: 38px">
<DIV style="MARGIN-TOP: 10px; FLOAT: left"><IMG 
src="../images/task_02.gif"></DIV>
<DIV style="MARGIN-TOP: 12px; FLOAT: left; MARGIN-LEFT: 10px"><A href="?sort=0"><SPAN class=anniu>ȫ���ȴ���������</SPAN></A>
  <A href="?sort=1"><SPAN class=anniu>�۸�����</SPAN></A>
  <A href="?sort=2"><SPAN class=anniu>�۵�����</SPAN></A>
  <A href="?sort=3"><SPAN class=anniu>1-30Ԫ��</SPAN></A>
  <A href="?sort=4"><SPAN class=anniu>31-100Ԫ��</SPAN></A>
  <A href="?sort=5"><SPAN class=anniu>101Ԫ������</SPAN></A></DIV>
<DIV style="CLEAR: right; MARGIN-TOP: 12px; FLOAT: right"><A title=���ˢ�� href="javascript:void(0)" onClick="reloadTask()" class="yell_font"> <SPAN class=anniu id=shuaimg>ˢ��ҳ��   ���񲻶�����</SPAN></A></DIV></DIV></DIV>

<div style="width:512px;height:32px;z_index:5px; display:none; position:fixed; _position:absolute;" id="waitingimage"><img src="../images/jieducm_shua.gif"/></div>
<div id="content">

<DIV style="CLEAR: both; WIDTH: 910px; BACKGROUND-COLOR: #f3f8fe"></DIV>
<DIV style="MARGIN-TOP: 0px; BACKGROUND: #ffffff; WIDTH: 910px">
<DIV style="CLEAR: both; WIDTH: 930px; BACKGROUND-COLOR: #ffffff">
<DIV style="CLEAR: both; WIDTH: 100%">
<DIV style="CLEAR: both; WIDTH: 100%; HEIGHT: 45px">
<DIV class=c_01 style="WIDTH: 15%"><IMG 
src="../images/task_01.gif">������</DIV>
<DIV class=c_01 style="WIDTH: 15%"><IMG 
src="../images/task_01.gif">������</DIV>
<DIV class=c_01 style="WIDTH: 15%"><IMG 
src="../images/task_01.gif">����۸�</DIV>
<DIV class=c_01 style="WIDTH: 15%"><IMG 
src="../images/task_01.gif">���۷�ʽ</DIV>
<DIV class=c_01 style="WIDTH: 20%"><IMG 
src="../images/task_01.gif">����״̬</DIV>
<DIV class=c_01 style="CLEAR: right; WIDTH: 20%"><IMG 
src="../images/task_01.gif">��&nbsp;&nbsp;��</DIV></DIV></DIV></DIV></DIV>
<DIV >
<%	 
action=request.QueryString("sort")
if action="" then
  Sql = "select top 100  * from "&jieducm&"_zhongxin where classid='2'  order by start asc,id desc"
elseif action=0 then
  Sql = "select top 100 * from "&jieducm&"_zhongxin where  classid='2' and start<>'5'  order by id desc, start  asc"
elseif action=1 then 
  Sql = "select top 100 * from "&jieducm&"_zhongxin where  classid='2' and  start='0' order by Price desc"
elseif action=2 then
  Sql = "select top 100 * from "&jieducm&"_zhongxin where   classid='2' and start='0' order by Price  asc"
elseif action=3 then
  Sql = "select top 100 * from "&jieducm&"_zhongxin  where  classid='2' and  price >= 1  and price<=30 and start='0' order by id desc"
elseif action=4 then
  Sql = "select top 100 * from "&jieducm&"_zhongxin where  classid='2' and  price >30  and price<=100 and start='0' order by id desc"
elseif action=5 then
  Sql = "select top 100 * from "&jieducm&"_zhongxin where   classid='2' and price >100  and start='0' order by id desc"
ElseIf action=6 Then
  Sql = "select top 100  * from "&jieducm&"_zhongxin where classid='2'  order by start asc,now desc"
end if
 Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then				
	if action="" then 
	  url="index.asp"
	else
	  url="index.asp?sort="&action&""
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
 bei=rs("bei")
 start=rs("start")
 pid=rs("pid")
 now1=rs("now")
 price=rs("price")
 isprice=rs("isprice")
 username=rs("username")
 tips=rs("tips")
 jieducm_fb=rs("jieducm_fb")
 addfabu=0
if bei="�����ջ�����"  then
clas=1
else
clas=2
end if
 %>  
<DIV class=tt5 <%if clas=2 then%> style="BACKGROUND-COLOR: #cde0ee" <%end if%>>
<TABLE style="MARGIN: 3px" cellSpacing=2 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD width="15%"><%=pid%><BR><%=now1%></TD>
    <TD>
    <TD align=middle width="15%"><SPAN  style="Z-INDEX: 20; ">
	 <%
			  	Sql2 = "select * from "&jieducm&"_register where username='"&rs("username")&"'"
				Set Rs2 = Server.CreateObject("Adodb.RecordSet")
				Rs2.Open Sql2,conn,1,1
				if rs2.eof then
				else	
				jifei=rs2("jifei")
				qq=rs2("qq")
			  %>
<% call isonline(username)%>&nbsp;<a href="../user/send_message.asp?sendname=<%=rs("username")%>" title="����վ����Ϣ" target="_blank"><%=rs("username")%></a>
<%end if%></SPAN><BR>
<img src="http://wpa.qq.com/pa?p=1:<%=qq%>:41"  border="0" />
<br />
<%call xinyu(jifei)%>
		 </TD>
    <TD align=middle width="15%">
	ƽ̨���� <img src="../images/jieducm_zf.gif" width="13" height="16"><font color="#FF0000"><strong><%=Price%></strong></font>Ԫ<br><%=rs("isprice")%></TD>
    <TD align=middle width="15%">
<%=bei%><BR><%
if bei="�����ջ�����"  then
response.Write("<IMG alt=�����ջ����� src=../images/jieducm_xuni.gif>")
else
response.Write("<IMG alt=�ӳ��ջ� src=../images/jieducm_shiwu.gif>")
end if%>������<font color="#FF0000"><strong><%
call jieducm_fadian(jieducm_fb,addfabu)
 %></strong></font>��</TD>
    <TD style="PADDING-LEFT: 50px" width="20%"><B   <%if start=0 then%> style="COLOR: red" <%else %>style="COLOR: #006600" <%end if%>>	
	<%call zhuangtai(start)%>
      </B><br /> <%if tips<>"" then
	response.Write("<br>(<a  title='"&tips&"'>")
	response.Write(left(tips,6))
	response.Write("...)")
	end if%></TD>
    <TD style="PADDING-LEFT: 50px" width="20%"> 
	<%
	if start="0" and rs("username")<>session("useridname")then
response.Write("<img src=../images/online_admin.gif>")%>
<a title="���֣����������ɻ�ô��ͷ�����" style="CURSOR: pointer" onClick="showxiao('<%=rs("pid")%>','2')">
<span class=anniu>��������</span></a>
<%elseif (start=1 or start=2 or start=3 or start=4) then
response.Write("�ѽ���")
elseif start=5 then
response.Write("�����ѽ���")
else
response.Write("���ܽ��Լ�������")
end if%>
</TD></TD></TR></TBODY></TABLE></DIV>
   <%
	  RowCount = RowCount - 1
	  i=i-1
      rs.MoveNext 
      Loop
 %> 
    <DIV class=tt5 >  <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"������") %></div>
  </div>

</DIV>

</div>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>